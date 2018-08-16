require 'rubyXL'
require './excel'

class Writer
  include Conf

  def write(summary)
    workbook = RubyXL::Parser.parse(out_path)
    worksheet = workbook[OUT_SHEET_NAME]
    worksheet[3][2].change_contents(2.4)
    count = 3
    #ファイル内の予実を取得
    until worksheet[count][NAME_ROW].value == nil do

      for i in 0..summary.count - 1

      end
      count = count + 1
    end

    # workbook.write(out_path)
    puts worksheet[3][2].value
  end

  def out_path
    File.join(OUT_FILE_PATH,OUTPUT_FILE)
  end

end
