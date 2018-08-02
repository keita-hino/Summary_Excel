require 'rubyXL'
# require 'win32ole'

class Excel
end

#コマンドライン引数を受け取る
PATH = ARGV[0]
FILE_TYPE = ".xlsx"

# input = "./target_folder/test.xlsx"
output = "output.xlsx"

Dir.glob("#{PATH}/*#{FILE_TYPE}").each do |input|
  puts input
  workbook = RubyXL::Parser.parse(input)
  worksheet = workbook['シート1']
  value = worksheet[3][2].value
  puts value
end

# エクセルファイルの読み込み
# workbook = RubyXL::Parser.parse(input)

# for i in 0..book.count
#   puts book[i]
# end

# buf[0][:"実績"] = 2.5
# puts buf[0][:"実績"]

# buf = [
#   "実績": 2.5,
#   "見込み": 4.5
# ]
# buf.push(
#   "実績": 5,
#   "見込み": 4
# )

buf = ["test": [
  "実績": 2.5,
  "見込み": 4.5
]]

# puts buf[0][:"test"][0][:"実績"]

# worksheet = workbook['シート1']
# value = worksheet[3][2].value
