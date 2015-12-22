class Student

  attr_reader :group, :id, :number, :name, :attend
  attr_accessor :score

  def initialize(group: '', id:, number: '', name: '', attend: true)
    @group = group
    @id = id
    @number = number
    @name = name
    @attend = attend
  end
end
