require "#{File.expand_path(File.dirname(__FILE__))}/constants"
require "#{File.expand_path(File.dirname(__FILE__))}/student"
require "#{File.expand_path(File.dirname(__FILE__))}/evaluation"
require 'roo'
require 'rubyXL'

class Main
  include Constants

  def initialize
    @reports_dir = REPORTS_DIR.sub(/\/\z/, '')

    @student_dirs = Dir.glob("#{@reports_dir}/**")
    @student_dirs.delete("#{@reports_dir}/reportlist.xls")
    @student_dirs.delete_if { |dir| dir =~ /evaluation_.+\.xlsx$/ }
    p "students: #{@student_dirs.size}"

    @students = []
    @evaluations = []
  end

  def check_evaluations
    # StudentとEvaluationをファイルから生成
    make_students
    make_evaluations

    # 真面目に評価しているか調べて、採点する
    mark_students_evaluation

    puts "Evaluation check has been completed."
  end

  def write_evaluations_in_class
    # クラス全体のファイルを書き出す
    evaluation_xlsx = RubyXL::Parser.parse("#{@reports_dir}/#{EVALUATION_DEFAULT_FILE_NAME}")

    @students.each do |student|
      row_num = evaluation_xlsx[EVALUATION_SHEET - 1].to_a.index { |row| row[EVALUATION_ID_COL - 1].value == student.id }
      col_num = EVALUATION_FIRST_ROW - 1

      @evaluations.select { |e| e.to_student == student }.each do |evaluation|
        unless evaluation.sougouten == '#DIV/0!'
          # TODO: sougoutenと指定のものになってしまっているのを直す
          evaluation_xlsx[EVALUATION_SHEET - 1][row_num][col_num].change_contents(evaluation.sougouten)
        end
        col_num += 1
      end

      evaluation_xlsx[EVALUATION_SHEET - 1][row_num][EVALUATION_FOR_OTHER_ROW - 1].change_contents(student.score)
    end

    write_file_name = "evaluation_#{Time.now.strftime('%Y%m%d%H%M%S')}.xlsx"
    evaluation_xlsx.write("#{@reports_dir}/#{write_file_name}")
    puts "Write #{write_file_name}"
  end

  def write_evaluations_for_student
    # 個人のフィードバックファイルを書き出す
    @students.each do |student|
      list_xlsx = RubyXL::Parser.parse(@reports_dir.sub(/reports$/, 'list.xlsx'))

      # TODO: 新しいファイルを作るようにする

      # # 左4列削除
      # delete_column_size = 4
      # delete_row_size = 55
      # delete_row_size.times do |row_num|
      #   delete_column_size.times do
      #     list_xlsx[0].delete_cell(row_num, 0, :left)
      #   end
      # end
      # # 列のサイズ調整
      # (delete_column_size + 1).times do |col_num|
      #   list_xlsx[0].change_column_width(col_num, 10.5)
      # end
      # list_xlsx[0].change_column_width(6, 30)

      # グループの人からの評価を書き込む（順番をランダムにする）
      from_students = @students.select do |s|
        (s.group == student.group) && s.attend
      end.shuffle
      from_students.delete(student)
      # TODO: 間違えたファイルを提出している(評価は無いが出席はしている)学生に対応する
      # 間違えたファイルを提出している人を削除
      from_students.delete_if { |s| s.id == 34 }

      # list_xlsx[0][12][0].change_contents('')
      from_students.each.with_index(1) do |from_student, i|
        row_num = 12 + i
        list_xlsx[0][row_num][4].change_contents("S#{i}")
        evaluation = @evaluations.find do |e|
          (e.from_student == from_student) && (e.to_student == student)
        end
        evaluation.make_array_evaluation_with_comment.each.with_index(5) do |value, col_num|
          list_xlsx[0][row_num][col_num].change_contents(value)
        end
      end

      list_xlsx.write("#{@reports_dir.sub(/reports$/, "lists/#{student.number}#{student.name}.xlsx")}")

      puts "Write #{student.number}#{student.name}.xlsx"
    end
  end

  private

  def make_students
    # 出席者を調べる
    attendances = @student_dirs.map { |student_dir| student_dir.sub(/\A#{@reports_dir}\/\d+-/, '') }

    first_student_num = @student_dirs.first.scan(/\/(\d+)$/).flatten.first
    xlsx_file = Roo::Excelx.new("#{@student_dirs.first}/#{FILE_PREFIX}#{first_student_num}.xlsx")
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

  def make_evaluations
    @student_dirs.each do |student_dir|
      student_number = student_dir.scan(/\/(\d+)$/).flatten.first
      xlsx_file = Roo::Excelx.new("#{student_dir}/#{FILE_PREFIX}#{dir_student_num}.xlsx")
      # 正しいファイルを提出しているかチェック
      # TODO: 正しいファイル(list.xlsx)と比較するように修正
      unless (['記入者', '採点項目', '点数'] - xlsx_file.sheet(STUDENT_SHEET - 1).row(1)).none?
        p "WARN: Incorrect file! - #{FILE_PREFIX}#{dir_student_num}.xlsx"
        next
      end

      from_student = @students.find { |student|  student.number == student_number }

      to_students = @students.select { |student| student.group == from_student.group }
      to_students.delete(from_student)

      # xlsx_fileから他のメンバーへの評価を取得
      to_students.each do |to_student|
        row_num = xlsx_file.sheet(STUDENT_SHEET - 1).column(STUDENT_ID_COL).index(to_student.id) + 1
        row = xlsx_file.sheet(STUDENT_SHEET - 1).row(row_num)

        # 正しい学生への評価か調べる
        unless row[STUDENT_GROUP_COL - 1] == to_student.group &&
               row[STUDENT_ID_COL - 1] == to_student.id &&
               row[STUDENT_NUMBER_COL - 1].to_s == to_student.number &&
               row[STUDENT_NAME_COL - 1] == to_student.name
          raise "Error: Incorrect student!! (student: #{to_student.name})"
        end

        evaluations = row[(STUDENT_EVALUATIONS_FIRST_COL - 1),(EVALUATIONS.size)]
        @evaluations << Evaluation.new(
          from: from_student,
          to: to_student,
          evaluations: evaluations
        )
      end
    end
  end

  def mark_students_evaluation
    @students.each do |student|
      evaluations = @evaluations.select do |evaluation|
        evaluation.from_student == student
      end

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
        student.score = 3
      end
    end
  end

  def all_evaluations_equal?(es)
    es.map(&:make_array_evaluation).uniq.size == 1
  end
end
