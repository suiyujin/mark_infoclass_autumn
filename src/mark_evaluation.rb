require "#{File.expand_path(File.dirname(__FILE__))}/constants"
require "#{File.expand_path(File.dirname(__FILE__))}/student"
require "#{File.expand_path(File.dirname(__FILE__))}/evaluation"
require 'roo'
require 'rubyXL'

class MarkEvaluation
  include Constants

  ### 初期化
  def initialize
    @reports_dir = REPORTS_DIR.sub(/\/\z/, '')

    @student_dirs = Dir.glob("#{@reports_dir}/**")
    @student_dirs.delete("#{@reports_dir}/reportlist.xls")
    @student_dirs.delete_if { |dir| dir =~ /#{EVALUATION_DEFAULT_FILE_NAME}$/ }
    @student_dirs.delete_if { |dir| dir =~ /evaluation_.+\.xlsx$/ }
    @student_dirs.delete_if { |dir| dir =~ /#{LIST_DEFAULT_FILE_NAME}$/ }
    @student_dirs.delete_if { |dir| dir =~ /lists$/ }
    puts "#{@student_dirs.size} studens."

    @students = []
    @evaluations = []
  end

  ### 情報をファイルから取得、採点
  def read_file_and_mark_evaluation
    # StudentとEvaluationをファイルから取得
    read_students
    read_evaluations

    # 真面目に評価しているか調べて、採点する
    mark_students_evaluation

    puts "Evaluation has been checked."
  end

  ### クラス全体の評価ファイルを書き出す
  def write_evaluations_in_class
    # テンプレートファイルを読み込む
    evaluation_xlsx = RubyXL::Parser.parse("#{@reports_dir}/#{EVALUATION_DEFAULT_FILE_NAME}")

    @students.each do |student|
      # 書き込む場所を指定
      row_index = evaluation_xlsx[EVALUATION_SHEET - 1].to_a.index { |row| row[EVALUATION_ID_COL - 1].value == student.id }
      col_index = EVALUATION_FIRST_COL - 1

      # 同じグループの学生からの評価(総合点)を書き込む
      @evaluations.select { |e| e.to_student == student }.each do |evaluation|
        unless evaluation.sougouten == '#DIV/0!'
          # TODO: sougoutenと指定のものになってしまっているのを直す
          evaluation_xlsx[EVALUATION_SHEET - 1][row_index][col_index].change_contents(evaluation.sougouten)
        end
        col_index += 1
      end

      # 他の人への評価に対する採点を書き込む
      evaluation_xlsx[EVALUATION_SHEET - 1][row_index][EVALUATION_FOR_OTHER_COL - 1].change_contents(student.score)
    end

    # ファイルを保存する
    write_file_name = "evaluation_#{Time.now.strftime('%Y%m%d%H%M%S')}.xlsx"
    evaluation_xlsx.write("#{@reports_dir}/#{write_file_name}")
    puts "Write #{write_file_name}"
  end

  ### 個人へのフィードバックファイルを書き出す
  def write_evaluations_for_students
    # lists/ディレクトリがなければ作成する
    unless File.exist?("#{@reports_dir}/lists/")
      FileUtils.mkdir("#{@reports_dir}/lists/")
      puts 'make directory: lists/'
    end

    @students.each do |student|
      # テンプレートファイルを読み込む
      list_xlsx = RubyXL::Parser.parse("#{@reports_dir}/#{LIST_DEFAULT_FILE_NAME}")

      # グループの人からの評価を用意（順番をランダムにする）
      from_students = @students.select do |s|
        (s.group == student.group) && s.attend
      end.shuffle
      from_students.delete(student)
      # TODO: 間違えたファイルを提出している(評価は無いが出席はしている)学生に対応する

      # 評価を書き込む
      from_students.each.with_index do |from_student, i|
        row_index = LIST_FIRST_ROW - 1 + i
        # 1列目に学生番号(S1,S2...)を書き込む
        list_xlsx[LIST_SHEET - 1][row_index][0].change_contents("S#{i}")

        # 評価を取得
        evaluation = @evaluations.find do |e|
          (e.from_student == from_student) && (e.to_student == student)
        end
        # 2列目以降に評価を書き込む
        evaluation.make_array_evaluation_with_comment.each.with_index(1) do |value, col_index|
          list_xlsx[LIST_SHEET - 1][row_index][col_index].change_contents(value)
        end
      end

      # ファイルを保存
      list_xlsx.write("#{@reports_dir}/lists/#{student.number}#{student.name}.xlsx")
      puts "Write lists/#{student.number}#{student.name}.xlsx"
    end
  end

  private

  ### 学生情報を取得
  def read_students
    # 出席者の学籍番号一覧をレポートのディレクトリから取得
    attendances = @student_dirs.map { |student_dir| student_dir.sub(/\A#{@reports_dir}\/\d+-/, '') }

    # 一番目の学生のレポートファイルを使う
    first_student_num = @student_dirs.first.scan(/-(\d+)\z/).flatten.first
    xlsx_file = Roo::Excelx.new("#{@student_dirs.first}/#{FILE_PREFIX}#{first_student_num}.xlsx")

    # 学生情報を取得してインスタンスを生成
    xlsx_file.each_row_streaming(pad_cells: true, offset: (STUDENT_LIST_FIRST_ROW - 2)) do |row|
      @students << Student.new(
        group: row[STUDENT_GROUP_COL - 1].value,
        id: row[STUDENT_ID_COL - 1].value,
        number: row[STUDENT_NUMBER_COL - 1].cell_value,
        name: row[STUDENT_NAME_COL - 1].value,
        attend: attendances.include?(row[STUDENT_NUMBER_COL - 1].cell_value)
      )
    end
  end

  ### 評価情報を取得
  def read_evaluations
    @student_dirs.each do |student_dir|
      dir_student_num = student_dir.scan(/-(\d+)\z/).flatten.first
      xlsx_file = Roo::Excelx.new("#{student_dir}/#{FILE_PREFIX}#{dir_student_num}.xlsx")
      # 正しいファイルを提出しているかチェック
      # TODO: 正しいファイル(list.xlsx)と比較するように修正
      unless (['記入者', '採点項目', '点数'] - xlsx_file.sheet(STUDENT_SHEET - 1).row(1)).none?
        p "WARN: Incorrect file! - #{FILE_PREFIX}#{dir_student_num}.xlsx"
        next
      end

      from_student = @students.find { |student| student.number == dir_student_num }
      to_students = @students.select { |student| student.group == from_student.group }
      to_students.delete(from_student)

      # ファイルから他のメンバーへの評価を取得
      to_students.each do |to_student|
        row_num = xlsx_file.sheet(STUDENT_SHEET - 1).column(STUDENT_ID_COL).index(to_student.id) + 1
        row = xlsx_file.sheet(STUDENT_SHEET - 1).row(row_num)

        # 正しい学生への評価か調べる
        unless row[STUDENT_GROUP_COL - 1] == to_student.group &&
               row[STUDENT_ID_COL - 1] == to_student.id &&
               row[STUDENT_NUMBER_COL - 1].to_s == to_student.number &&
               row[STUDENT_NAME_COL - 1] == to_student.name
          raise StandardError, "Incorrect student!! (student: #{to_student.name})"
        end

        # 評価インスタンスを生成
        evaluations = row[(STUDENT_EVALUATIONS_FIRST_COL - 1),(EVALUATIONS.size)]
        @evaluations << Evaluation.new(
          from: from_student,
          to: to_student,
          evaluations: evaluations
        )
      end
    end
  end

  ### 他の人への評価を採点する
  ### 0〜3で採点
  def mark_students_evaluation
    @students.each do |student|
      evaluations = @evaluations.select do |evaluation|
        evaluation.from_student == student
      end

      # 評価に応じて点数を付ける
      if evaluations.empty? ||
         evaluations.map(&:not_exist_all_evaluation?).all?
        # ファイルが提出されていない
        # ファイルは提出されているが、全ての評価が空
        student.score = 0
      elsif evaluations.map(&:all_num_equal?).all? ||
            all_evaluations_equal?(evaluations)
        # 各学生への評価ごとに全て同じ数字を使っている
        # 各学生への評価が数字の使い回しである
        student.score = 1
      elsif evaluations.map { |e| e.exist_nil_evaluation? && e.to_student_attend? }.any?
        # 評価に1つでも抜けがある
        # (欠席者への評価は除く)
        student.score = 2
      else
        # 上記以外
        student.score = 3
      end
    end
  end

  ### 各学生への評価が数字の使い回しかどうか調べる
  ### 使い回しの場合はtrueを返す
  def all_evaluations_equal?(es)
    es.map(&:make_array_evaluation).uniq.size == 1
  end
end
