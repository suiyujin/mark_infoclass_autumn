class Evaluation

  attr_reader :from_student, :to_student, :naiyo, :shiryohyogen, :hanashikata,
    :shinko, :doryokudo, :comment, :total

  def initialize(naiyo, shiryohyogen, hanashikata, shinko, doryokudo, comment, total, from:, to:)
    @from_student = from
    @to_student = to

    @naiyo = naiyo
    @shiryohyogen = shiryohyogen
    @hanashikata = hanashikata
    @shinko = shinko
    @doryokudo = doryokudo
    @comment = comment

    @total = total
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
    [@naiyo, @shiryohyogen, @hanashikata, @shinko, @doryokudo]
  end
end
