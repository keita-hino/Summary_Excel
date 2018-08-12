require 'rubyXL'
require './Config_Reader'

#コマンドライン引数を受け取る
# PATH = ARGV[0]
#本来は上記だが、開発ではこっち
PATH        = "/Users/hinokeita/Desktop/workspace/summary_excel/target_folder".freeze
CONFIG_FILE = "env.rb".freeze
OUTPUT_FILE = "output.xlsx".freeze
FILE_TYPE   = ".xlsx".freeze
SHEET_NAME  = "シート1".freeze

sum_array = [{}]

#IDと名前の読み取り
config = ConfigReader.new

config.instance_eval File.read CONFIG_FILE

buf = config.read
buf.each do |i|
  # puts i[:"id"]
  # puts i[:"name"]
end
puts buf

# input = "./target_folder/test.xlsx"

#フォルダ内の全ファイル取得
Dir.glob("#{PATH}/*#{FILE_TYPE}").each do |input|
  count = 3
  # puts input
  workbook = RubyXL::Parser.parse(input)
  worksheet = workbook[SHEET_NAME]

  #ファイル内の予実を取得
  while !(worksheet[count][1].value == nil) do
    pj_name = worksheet[count][1].value
    zisseki =  worksheet[count][2].value
    mikomi = worksheet[count][3].value
    sum_array.push(
      "#{pj_name}":[
      "実績":zisseki,
      "見込み":mikomi
    ])
    count = count + 1
  end
end

#配列を初期化した際の要素を削除
sum_array.delete_at(0)
puts sum_array

#提出してもらったエクセルファイルをハッシュに格納する処理
# buf = ["test": [
#   "実績": 2.5,
#   "見込み": 4.5
# ]]

# puts buf[0][:"test"][0][:"実績"]
