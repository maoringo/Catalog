module ExcelBase
    require 'rubygems'
    require 'spreadsheet'
    
    Spreadsheet.client_encoding = 'UTF-8'

    class NormalFormat < Spreadsheet::Format
        def initialize(opts={})
            super opts.merge(:size => 11)
            super opts.merge(:name => 'ヒラギノ丸ゴ Pro') 
        end
    end

    class ColorFormat < NormalFormat
        def initialize(color)
            super :pattern => 1, :pattern_fg_color => color
        end
    end
    
end
