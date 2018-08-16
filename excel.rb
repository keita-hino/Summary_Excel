require 'rubyXL'
require './Config_Reader'


module Conf
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
  include Conf

  def read
    #IDと名前の読み取り
    config = conf_read

    reader = Reader.new
    reader.read(config)
  end

  def conf_read
    config = ConfigReader.new
    config.instance_eval File.read CONFIG_FILE
    # config.read
  end

  def write(summary)
    writer = Writer.new
    writer.write(summary)
  end

end
