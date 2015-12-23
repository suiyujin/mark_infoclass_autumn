require "#{File.expand_path(File.dirname(__FILE__))}/mark_evaluation"

# 初期化処理
mark_evaluation = MarkEvaluation.new

# 情報をファイルから取得、採点
mark_evaluation.read_file_and_mark_evaluation

# クラス全体の評価ファイルを書き出す
mark_evaluation.write_evaluations_in_class

# 個人へのフィードバックファイルを書き出す
mark_evaluation.write_evaluations_for_students

puts "all execute is complete."
