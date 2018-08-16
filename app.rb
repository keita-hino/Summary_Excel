require './excel'
require 'json'
require './reader'
require './writer'

excel = Excel.new
summary = excel.read
# puts summary
# puts summary[0][:"Sample3"]

excel.write(summary)

# puts a

# excel.write
