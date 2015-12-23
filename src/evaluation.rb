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

  def to_student_attend?
    @to_student.attend
  end

  def exist_nil_evaluation?
    make_array_evaluation_with_comment.include?(nil)
  end

  def not_exist_all_evaluation?
    make_array_evaluation_with_comment.compact.empty?
  end

  def all_num_equal?
    make_array_evaluation.uniq.size == 1
  end

  def make_array_evaluation_with_comment
    make_array_evaluation.push(@comment)
  end

  def make_array_evaluation
    ary = EVALUATIONS.map { |evaluation| instance_variable_get("@#{evaluation}") }
    ary[0, 5]
  end
end
