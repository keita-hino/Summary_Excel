require 'rubyXL'

class Excel
  include RubyXL
  def find_cell
    puts "out"
  end
end

excel = Excel.new
# puts Excel::Parser.parse("/Users/hinokeita/Desktop/workspace/summary_excel/target_folder/test1.xlsx")
# puts Excel::Wrokshees.each

#コマンドライン引数を受け取る
PATH = ARGV[0]
FILE_TYPE = ".xlsx"
SHEET_NAME = "シート1"
sum_array = [{}]

# input = "./target_folder/test.xlsx"
output = "output.xlsx"

#フォルダ内の全ファイル取得
Dir.glob("#{PATH}/*#{FILE_TYPE}").each do |input|
  count = 3
  puts input
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
