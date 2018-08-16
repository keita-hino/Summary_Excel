require 'rubyXL'
require './excel'

class Reader
  include Conf

  def read(config)
    sum_array = [{}]

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

end
