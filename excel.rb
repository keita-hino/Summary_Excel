require 'rubyXL'
require './Config_Reader'

module CONF
  IN_FILE_PATH   = "/Users/hinokeita/Desktop/workspace/summary_excel/target_folder".freeze
  OUT_FILE_PATH  = "/Users/hinokeita/Desktop/workspace/summary_excel/output_file".freeze
  CONFIG_FILE    = "env.rb".freeze
  OUTPUT_FILE    = "output.xlsx".freeze
  FILE_TYPE      = ".xlsx".freeze
  IN_SHEET_NAME  = "シート1".freeze
  OUT_SHEET_NAME = "20180811".freeze
  GID_ROW        = 1
  PJNAME_ROW     = 2
  ZISSEKI        = 3
  MIKOMI         = 4
  NAME_ROW       = 1
end

class Excel
  include CONF

  def read
    sum_array = [{}]

    #IDと名前の読み取り
    config = conf_read

    #フォルダ内の全ファイル取得
    Dir.glob("#{IN_FILE_PATH}/*#{FILE_TYPE}").each do |input|
      count = 3
      workbook = RubyXL::Parser.parse(input)
      worksheet = workbook[IN_SHEET_NAME]

      #ファイル内の予実を取得
      until worksheet[count][GID_ROW].value == nil do
        name = convert_gid(config,worksheet[count][GID_ROW].value.to_s)
        pj_name = worksheet[count][PJNAME_ROW].value
        zisseki = worksheet[count][ZISSEKI].value
        mikomi = worksheet[count][MIKOMI].value
        sum_array.push(
          "#{name}":[
            "#{pj_name}":[
              "実績":zisseki,
              "見込み":mikomi
            ]
          ]
        )
        count += 1
      end
    end

    #配列を初期化した際の要素を削除
    sum_array.delete_at(0)
    return sum_array
  end

  def convert_gid(config,gid)
    for i in 0..config.count - 1
      unless config[i][:"#{gid}"].nil?
        return config[i][:"#{gid}"]
        break
      end
    end
  end

  def conf_read
    config = ConfigReader.new
    config.instance_eval File.read CONFIG_FILE
    # config.read
  end

  def write(summary)
    workbook = RubyXL::Parser.parse(out_path)
    worksheet = workbook[OUT_SHEET_NAME]
    worksheet[3][2].change_contents(2.4)
    count = 3
    # puts summary[0][:"Sample3"]
    #ファイル内の予実を取得
    until worksheet[count][NAME_ROW].value == nil do

      for i in 0..summary.count - 1

      end
      puts worksheet[count][NAME_ROW].value
      count = count + 1
    end

    # workbook.write(out_path)
    puts worksheet[3][2].value
  end

  def out_path
    File.join(OUT_FILE_PATH,OUTPUT_FILE)
  end

end
