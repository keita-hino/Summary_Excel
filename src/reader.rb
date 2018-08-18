require 'rubyXL'
require './excel'
require 'byebug' #デバッグ用

class Reader
  include Conf

  def initialize
    @count
    @workbook
    @worksheet
  end

  def read(config)
    sum_array = [{}]

    #フォルダ内の全ファイル取得
    Dir.glob("#{IN_FILE_PATH}/*#{FILE_TYPE}").each do |input|
      @count = 3
      get_workbook(input)
      #ファイル内の予実を取得
      until worksheet[@count][GID_ROW].value == nil do
        sum_array.push(
          "#{name(config)}":[
            "#{pj_name}":[
              "実績":zisseki,
              "見込み":mikomi
            ]
          ]
        )
        @count += 1
      end
    end

    #配列を初期化した際の要素を削除
    sum_array.delete_at(0)
    return sum_array
  end

  def get_workbook(input)
    @workbook =  RubyXL::Parser.parse(input)
  end

  def worksheet
    @worksheet = @workbook[IN_SHEET_NAME]
  end

  def name(config)
    convert_gid(config,@worksheet[@count][GID_ROW].value.to_s)
  end

  def pj_name
    @worksheet[@count][PJNAME_ROW].value
  end

  def zisseki
    @worksheet[@count][ZISSEKI].value
  end

  def mikomi
    @worksheet[@count][MIKOMI].value
  end

  def convert_gid(config,gid)
    for i in 0..config.count - 1
      unless config[i][:"#{gid}"].nil?
        return config[i][:"#{gid}"]
        break
      end
    end
  end

end
