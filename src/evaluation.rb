class Evaluation
  include Constants

  attr_reader :from_student, :to_student

  def initialize(from:, to:, evaluations:)
    @from_student = from
    @to_student = to

    # configファイルの評価一覧からインスタンス変数とアクセサメソッドを作成
    EVALUATIONS.each_with_index do |evaluation, i|
      instance_variable_set("@#{evaluation}", evaluations[i])
    end
    class << self
      EVALUATIONS.each do |evaluation|
        define_method("#{evaluation}") { eval("@#{evaluation}") }
      end
    end
  end

  # 評価されている学生が出席しているか
  def to_student_attend?
    @to_student.attend
  end

  # 評価に一つでも抜けがあるか
  def exist_nil_evaluation?
    make_array_evaluation_with_comment.include?(nil)
  end

  # 全ての評価が空か
  def not_exist_all_evaluation?
    make_array_evaluation_with_comment.compact.empty?
  end

  # 全て同じ数字で評価されているか
  def all_num_equal?
    make_array_evaluation.uniq.size == 1
  end

  # コメント付き評価配列を生成
  def make_array_evaluation_with_comment
    make_array_evaluation.push(@comment)
  end

  # コメント無し評価配列を生成
  def make_array_evaluation
    ary = EVALUATIONS.map { |evaluation| instance_variable_get("@#{evaluation}") }
    ary[0, 5]
  end
end
