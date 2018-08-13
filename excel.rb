require 'rubyXL'
require './Config_Reader'

class Excel

  #コマンドライン引数を受け取る
  # IN_FILE_PATH = ARGV[0]
  #本来は上記だが、開発ではこっち
  IN_FILE_PATH   = "/Users/hinokeita/Desktop/workspace/summary_excel/target_folder".freeze
  OUT_FILE_PATH  = "/Users/hinokeita/Desktop/workspace/summary_excel/output_file".freeze
  CONFIG_FILE    = "env.rb".freeze
  OUTPUT_FILE    = "output.xlsx".freeze
  FILE_TYPE      = ".xlsx".freeze
  IN_SHEET_NAME  = "シート1".freeze
  OUT_SHEET_NAME = "20180811".freeze

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
      while !(worksheet[count][2].value == nil) do
        name = convert_gid(config,worksheet[count][1].value.to_s)
        pj_name = worksheet[count][2].value
        zisseki = worksheet[count][3].value
        mikomi = worksheet[count][4].value
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
    return sum_array
    # puts sum_array

    # puts buf[0][:"test"][0][:"実績"]
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

  def write
    workbook = RubyXL::Parser.parse(out_path)
    worksheet = workbook[OUT_SHEET_NAME]
    worksheet[3][2].change_contents(2.4)


    # workbook.write(out_path)
    puts worksheet[3][2].value
  end

  def out_path
    File.join(OUT_FILE_PATH,OUTPUT_FILE)
  end

end
