require 'roo'
require 'spreadsheet'
require 'kconv'
require "csv"
require './excel_base'
include ExcelBase

catalog_ary = []
fromCatalogAry = []
#カタログデータの読み込み
CSV.open("integbio_dbcatalog_20131030_utf8.csv",'r'){|row|
catalog_ary =  row.to_a
}
catalog_ary.each do |elem|
fromCatalogAry << elem[1]
end

#もとの調査ファイルの読み込み
lineAry = []
lineArys = []
spst = Roo::Excel.new(ARGV[0])
spst.default_sheet = spst.sheets.first
1.upto(spst.last_row) do |line|
for i in 1..17
lineArys << spst.cell(line,i).to_s
end
lineAry << lineArys
lineArys = []
end
 
default_format = NormalFormat.new
workbook = Spreadsheet::Workbook.new
workbook.default_format = default_format
worksheet = workbook.create_worksheet

#該当するものを色付け
lineAry.each_with_index do |line,i|
    line.each_with_index do |elem,j|
        worksheet[i,j] = elem
            if elem.include?("データベース")
            worksheet.row(i).set_format(j,ColorFormat.new(:orange))
            end
            fromCatalogAry.each_with_index do |celem,k|
             if elem.include?(celem)
                worksheet.row(i).set_format(j,ColorFormat.new(:blue))
             end
            end
    end
end

workbook.write('sample2.xls')
