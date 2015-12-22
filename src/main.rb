require "#{File.expand_path(File.dirname(__FILE__))}/student.rb"
require "#{File.expand_path(File.dirname(__FILE__))}/evaluation.rb"
require 'roo'
require 'rubyXL'

class Main

  def initialize
    @reports_dir = File.expand_path(File.dirname(__FILE__)).sub(/src/, 'reports')
    @file_prefix = 'kadai07-'
    @student_dirs = Dir.glob("#{@reports_dir}/**")
    @student_dirs.delete("#{@reports_dir}/reportlist.xls")
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
    evaluation_xlsx = RubyXL::Parser.parse(@reports_dir.sub(/reports$/, 'evaluation_default.xlsx'))

    @students.each do |student|
      row_num = evaluation_xlsx[0].to_a.index { |row| row[1].value == student.id }
      col_num = 4

      @evaluations.select { |e| e.to_student == student }.each do |evaluation|
        evaluation_xlsx[0][row_num][col_num].change_contents(evaluation.total)
        col_num += 1
      end

      col_num = 11
      evaluation_xlsx[0][row_num][col_num].change_contents(student.score)
    end

    evaluation_xlsx.write(@reports_dir.sub(/reports$/, 'evaluation.xlsx'))
    puts "Write evaluation.xlsx"
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
    attendances = Dir.entries(@reports_dir) - ['.', '..', 'reportlist.xls']

    first_student_num = @student_dirs.first.scan(/\/(\d+)$/).flatten.first
    xlsx_file = Roo::Excelx.new("#{@student_dirs.first}/#{@file_prefix}#{first_student_num}.xlsx")
    xlsx_file.each_row_streaming(pad_cells: true, offset: 12) do |row|
      @students << Student.new(
        group: row[0].value,
        id: row[1].value,
        number: row[2].cell_value,
        name: row[3].value,
        attend: attendances.include?(row[2].cell_value)
      )
    end
  end

  def make_evaluations
    @student_dirs.each do |student_dir|
      student_number = student_dir.scan(/\/(\d+)$/).flatten.first
      xlsx_file = Roo::Excelx.new("#{student_dir}/#{@file_prefix}#{student_number}.xlsx")
      # 正しいファイルを提出しているかチェック
      unless (['記入者', '採点項目', '点数'] - xlsx_file.sheet(0).row(1)).none?
        p "WARN: Incorrect file! - #{@file_prefix}#{student_number}.xlsx"
        next
      end

      from_student = @students.find { |student|  student.number == student_number }

      to_students = @students.select { |student| student.group == from_student.group }
      to_students.delete(from_student)

      # xlsx_fileから他のメンバーへの評価を取得
      to_students.each do |to_student|
        row_num = xlsx_file.sheet(0).column(2).index(to_student.id) + 1
        row = xlsx_file.sheet(0).row(row_num)

        # 正しい学生への評価か調べる
        unless row[0] == to_student.group &&
               row[1] == to_student.id &&
               row[2].to_s == to_student.number &&
               row[3] == to_student.name
          raise "Error: Incorrect student!! (student: #{to_student.name})"
        end

        evaluations = row[5..-2]
        @evaluations << Evaluation.new(
          *evaluations,
          from: from_student,
          to: to_student
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
