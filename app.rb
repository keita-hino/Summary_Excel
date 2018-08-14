require './excel'
require 'json'

excel = Excel.new
summary = excel.read
puts summary[0][:"Sample3"]


# excel.write(summary)

# puts a

# excel.write
