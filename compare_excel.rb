require 'creek'
require 'write_xlsx'

puts "START GET DATA SHEET"
creek = Creek::Book.new 'Contact Rikai.xlsx'
sheet_1 = creek.sheets[0]
sheet_2 = creek.sheets[1]

data_url_sheet_1 =[]
data_sheet_2 =[]

sheet_1.rows.each_with_index do |row,index|
    next if row.values[3].nil? || index == 0
    data_url_sheet_1 << row.values[3]
end

sheet_2.rows.each_with_index do |row,index|
    next if row.values[3].nil? || index == 0
    data_sheet_2 << row.values
end

puts "END GET DATA SHEET"
puts "===================="
puts "START COMPARE"

workbook = WriteXLSX.new('Convert.xlsx')
worksheet = workbook.add_worksheet

data_url_sheet_2 = []
data_sheet_2.each_with_index do |row,index|
    data_url_sheet_2 << row[3]
end


data = data_url_sheet_1 + data_url_sheet_2
data_not_dupplicate = data - (data_url_sheet_1 & data_url_sheet_2)
data_new = data_url_sheet_2 & data_not_dupplicate

array_new = []

worksheet.write(0,   0, "Key")
worksheet.write(0,   1, "Company Name")
worksheet.write(0,   2, "Category")
worksheet.write(0,   3, "Category url")

count = 1
data_sheet_2.each_with_index do |row, index|
    if (row & data_new).length > 0
        worksheet.write(count,   0, row[0])
        worksheet.write(count,   1, row[1])
        worksheet.write(count,   2, row[2])
        worksheet.write(count,   3, row[3])
        count += 1
    end
    puts "IS READING LINE: #{index}"
end

puts "COMPARED IS SUCCESSFUL"
puts "DONE"
workbook.close
