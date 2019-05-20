require 'axlsx'

require 'rexml/document'
include REXML



class AXLSX_create

    def show_grid_lines(sheet, doc)
        if REXML::XPath.first(doc,'//DoNotDisplayGridlines')
            grid_lines_status = sheet.sheet_view.show_grid_lines = false
        else 
            grid_lines_status = sheet.sheet_view.show_grid_lines = true
        end
        return grid_lines_status
    end

    def get_alignment(doc)
        # define your styles
        #title = ws.styles.add_style(:bg_color => "FFFF0000", :fg_color=>"#FF000000", :border=>Axlsx::STYLE_THIN_BORDER,:alignment=>{:horizontal => :center})
        alignment={} 
        unless REXML::XPath.first(doc,'Alignment').nil?
            alignment[:horizontal]=REXML::XPath.first(doc,'Alignment').get_attribute("ss:Horizontal").value unless 
                REXML::XPath.first(doc,'Alignment').get_attribute("ss:Horizontal").nil?
            alignment[:vertical]=REXML::XPath.first(doc,'Alignment').get_attribute("ss:Vertical").value unless 
                REXML::XPath.first(doc,'Alignment').get_attribute("ss:Vertical").nil?
        end
        return alignment # à tester si c'est nul
    end

    def set_style_alignment(doc)
        alignment = get_alignment(doc)
        if alignment[:horizontal].nil?
            style_alignment = ':alignment=>{:vertical => :' + alignment[:vertical].to_s.downcase + '}'
        elsif alignment[:vertical].nil?
            style_alignment = ':alignment=>{:horizontal => :' + alignment[:horizontal].to_s.downcase + '}'
        else
            style_alignment = ':alignment=>{:horizontal => :' + alignment[:horizontal].to_s.downcase + 
                ', :vertical => :'+ alignment[:vertical].to_s.downcase  + '}'
        end
        return style_alignment
    end

    def get_text_color(element)
        return  element.elements["Font"].attributes["ss:Color"].gsub('#','') unless 
                    element.elements["Font"].nil? || element.elements["Font"].attributes["ss:Color"].nil?
    end

    def set_text_color(element)
        return text_color = get_text_color(element) unless get_text_color(element).nil?
    end

    def get_text_size(element)
        return  element.elements["Font"].attributes["ss:Size"] unless 
                    element.elements["Font"].nil? || element.elements["Font"].attributes["ss:Size"].nil?
    end

    def set_text_size(element)
        return text_size =  get_text_size(element).to_i unless get_text_size(element).nil?
    end

    def get_text_bold(element)
        return  element.elements["Font"].attributes["ss:Bold"] unless 
                    element.elements["Font"].nil? || element.elements["Font"].attributes["ss:Bold"].nil?
    end

    def set_text_bold(element)
        return text_bold = get_text_bold(element) == "1" ? true : false unless get_text_bold(element).nil?
    end

    def get_text_italicised(element)
        return  element.elements["Font"].attributes["ss:Italic"] unless
                    element.elements["Font"].nil? || element.elements["Font"].attributes["ss:Italic"].nil?
    end

    def set_text_italicised(element)
        return text_italic = get_text_italicised(element) == "1" ? true : false unless get_text_italicised(element).nil?
    end

    def get_text_underlined(element)
        return  element.elements["Font"].attributes["ss:Underline"] unless 
                element.elements["Font"].nil? || element.elements["Font"].attributes["ss:Underline"].nil?    
    end

    def set_text_underlined(element)
        return text_underline = get_text_underlined(element) == "1" ? true : false  unless get_text_underlined(element).nil?
    end

    def get_font_name(element)
        return  element.elements["Font"].attributes["ss:FontName"] unless 
                element.elements["Font"].nil? || element.elements["Font"].attributes["ss:FontName"].nil?    
    end

    def set_font_name(element)
        return font_name = get_font_name(element) unless get_font_name(element).nil?
    end

    def get_bg_color(element)
        return  element.elements["Interior"].attributes["ss:Color"].gsub('#','') unless 
                element.elements["Interior"].nil? || element.elements["Interior"].attributes["ss:Color"].nil?    
    end

    def set_bg_color(element)
        return bg_color = get_bg_color(element) unless get_bg_color(element).nil?
    end

    def get_format_cell(element)
        return  element.elements["NumberFormat"].attributes["ss:Format"] unless
                    element.elements["NumberFormat"].nil? || element.elements["NumberFormat"].attributes["ss:Format"].nil?     
    end

    def set_format_cell(element)
        return format_cell = get_format_cell(element).to_sym unless get_format_cell(element).nil?
    end

    def get_border_type_cell(element)
        border_type=""
         #unless element.nil? || element.attributes["ss:LineStyle"].nil?
            if element.attributes["ss:LineStyle"] =="Continuous"
                if element.attributes["ss:Weight"] == 2
                    border_type = "medium"
                else
                    border_type = "thick"
                end
            elsif element.attributes["ss:LineStyle"] =="Dash"
                if element.attributes["ss:Weight"] == 1
                    border_type = "dashed"
                else
                    border_type = "mediumDashed"
                end
            elsif element.attributes["ss:LineStyle"] =="DashDot"
                if element.attributes["ss:Weight"] == 1
                    border_type = "dashDot"
                else
                    border_type = "mediumDashDot"
                end
            elsif element.attributes["ss:LineStyle"] =="DashDotDot"
                if element.attributes["ss:Weight"] == 1
                    border_type = "dashDotDot"
                else
                    border_type = "mediumDashDotDot"
                end
            elsif element.attributes["ss:LineStyle"] =="Dot"
                border_type = "dotted"
            elsif element.attributes["ss:LineStyle"] =="SlantDashDot"
                border_type = "slantDashDot"
            elsif element.attributes["ss:LineStyle"] =="Double"
                border_type = "double"
            else
                border_type = "thin"
            end 
        #end           
        return border_type.to_sym    
    end

    def set_border_type_cell(element)
        return  border_cell = get_border_type_cell(element) unless get_border_type_cell(element).nil?
    end

    def get_border_color_cell(element)
        return  element.attributes["ss:Color"].gsub('#','') unless
                    element.nil? || element.attributes["ss:Color"].nil?        
    end

    def set_border_color_cell(element)
        return format_cell = get_border_color_cell(element) unless get_border_color_cell(element).nil?
    end

    def get_border_position(element)
        return  element.attributes["ss:Position"] unless
                    element.nil? || element.attributes["ss:Position"].nil?        
    end

    def set_border_position(element)
        return format_cell = get_border_position(element).downcase.to_sym unless get_border_position(element).nil?
    end

    
    def set_border_cell(element)
        border_cell = {}
        border_cell[:style] = set_border_type_cell(element) unless element.nil? ||  set_border_type_cell(element).nil?
        border_cell[:color] = set_border_color_cell(element) unless element.nil? ||  set_border_color_cell(element).nil?
        border_cell[:edges] = [set_border_position(element)] unless element.nil? ||  set_border_position(element).nil?
        return border_cell
    end


    def get_style(element)
        styles = {}
        puts element.attributes["ss:ID"]
        styles[:id] = element.attributes["ss:ID"] || "toto"
        styles[:b] = set_text_bold(element) unless set_text_bold(element).nil? || element.elements["Font"].attributes["ss:Bold"].nil?
        styles[:i] = set_text_italicised(element) unless set_text_italicised(element).nil? || element.elements["Font"].attributes["ss:Italic"].nil?
        styles[:u] = set_text_underlined(element) unless set_text_underlined(element).nil? || element.elements["Font"].attributes["ss:Underline"].nil?
        styles[:sz] = set_text_size(element) unless set_text_size(element).nil? || element.elements["Font"].attributes["ss:Size"].nil?
        styles[:fg_color] = set_text_color(element) unless set_text_color(element).nil? || element.elements["Font"].attributes["ss:Color"].nil?
        styles[:font_name] = set_font_name(element) unless set_font_name(element).nil? || element.elements["Font"].attributes["ss:FontName"].nil?
        styles[:bg_color] = set_bg_color(element) unless set_bg_color(element).nil? || element.elements["Interior"].attributes["ss:Color"].nil?
        #styles[:format_code] = set_format_cell(element) unless set_format_cell(element).nil? || element.elements["NumberFormat"].attributes["ss:Format"].nil?
        element.get_elements("Borders").each{ |el|
            styles[:border]= [] unless  el.get_elements("Border").nil?
            el.get_elements("Border").each{ |e|
                styles[:border].push(set_border_cell(e))
            }
        } unless  element.get_elements("Borders").nil?
        return styles
    end

    def set_style(value_id, array_of_hashes)
        res ={}
        array_of_hashes.any? {|h|
            puts h 
            if h[:id] == value_id
                h.delete(:id)
                h.delete(:border)
                 res = h.clone
            end
        }
        return res
    end
    def set_style_border(value_id, array_of_hashes)
        res ={}
        gridstyle_border2
        array_of_hashes.any? {|h|
            if h[:id] == value_id
                 h[:border].each{ |x|
                    gridstyle_border2= s.add_style :border => x
                }
            end
        }
        return gridstyle_border2
    end

    def get_column_position_first_cell_in_row(element)
        return element.attributes["ss:Index"] unless 
                element.nil? || element.attributes["ss:Index"].nil?
    end
    
    def get_value_cell(element)
        return element.elements["Data"].get_text.value  unless element.elements["Data"].nil?
    end

    def get_row_value(element, rows, default_style)
        row_value = {}
        num_col = 1
        a = element.to_s.gsub(/\r/,'').gsub(/\t/,'').gsub(/\n/,'').tr("'", '"')
        res = a.gsub("ss:",' ')
        #puts res 
        doc = REXML::Document.new(res)
        at =  []
        i = 0
        col = 0
        column = 1
        style=[]
        doc.elements["Row"].elements.each('Cell') { |element|
            if element.attributes["Index"].nil?
                col = col + 1
            else 
                col = element.attributes["Index"].to_i
            end
            while (column < col)
                at.push("")
                style.push(default_style)
                column = column + 1
            end
            if column == col
                if element.attributes["StyleID"].nil?
                    style.push(default_style)
                else
                    style.push(element.attributes["StyleID"].to_s)
                end
                column = column + 1
                puts "colonne #{col}"
                data=element.elements['Data'].get_text unless element.elements['Data'].nil?
                at.push(data.to_s)
            end
        }
        #puts "le style est #{style}"
        #puts "la row est #{rows} "
        #puts "les values sont #{at}"
        #puts "######################################################"
        return [rows, at, style]
    end

    def get_default_style(element)
        return  element.attributes["ss:StyleID"] unless 
                element.nil? || element.attributes["ss:StyleID"].nil? 
    end

    def get_tab_style(c, array_style, s)
        res=[]
        puts "WHAAAAAAAAAAAAAAAAT #{array_style}"
        c.each { |x| 
            array_style.find {|el| el[:id].to_s === x.to_s
                #puts "WHAAAAAAAAAAAAAAAAT #{el}"
                el.delete(:id)
                gridstyle_border = s.add_style(el) unless el.nil?
                #res.push(gridstyle_border)
            }
            #puts "WHAAAAAAAAAAAAAAAAT #{el}"            
        }
        return res
    end


end

package = Axlsx::Package.new
doc = REXML::Document.new(File.read('#########################################.xml'))
test = AXLSX_create.new 
array_style = []
default_style=""
doc.elements.each("Workbook/Styles/Style") { |element| 
    array_style.push(test.get_style(element))
}
puts "AAAAAAAAAAAAAAA #{array_style}"
doc.elements.each("Workbook/Worksheet/Table") { |element| 
    nb_columns = element.attributes["ss:ExpandedColumnCount"]
    nb_rows = element.attributes["ss:ExpandedRowCount"]
    puts "nb columns = #{nb_columns} and nb rows = #{nb_rows} "
    default_style = test.get_default_style(element) || "Default"
}
rows = 0
current_row = 1
testt = []
package.workbook do |workbook|
    workbook.styles do |s|
        gridstyle_border =  s.add_style :border => { :style => :hair, :color =>"FFCDCDCD", :edges => [:top, :bottom] }
        gridstyle_border2 =  s.add_style :border => { :style => :hair, :color =>"FFC92A2A", :edges => [:top, :bottom] }
        testt.push(gridstyle_border)
        testt.push(gridstyle_border2)
        puts "BOUYAG#{test}"
        workbook.add_worksheet :name => "Custom Borders"  do |sheet|
            doc.elements.each("Workbook/Worksheet/Table/Row") { |element|
                if test.get_column_position_first_cell_in_row(element).nil?
                    rows = rows.to_i + 1
                else
                    rows = test.get_column_position_first_cell_in_row(element).to_i
                end 
                test.show_grid_lines(sheet, doc)
                a, b, c = test.get_row_value(element,rows,default_style)
                d = test.get_tab_style(c, array_style, s)
                while (current_row < a)
                    current_row = current_row + 1
                    #puts c
                    sheet.add_row [], :style => d || []
                end
                sheet.add_row b, :style => d || []
                current_row = current_row + 1
            }
        end
    end
end

#b = test.set_style(value_id, array_style)
#package.workbook do |workbook|
  #workbook.styles do |s|
    #gridstyle_border =  s.add_style :border => { :style => :hair, :color =>"FFCDCDCD", :edges => [:top, :bottom] }
    #style_test = s.add_style(b)
    #workbook.add_worksheet :name => "Custom Borders"  do |sheet|
      #test.show_grid_lines(sheet, doc)
      #sheet.add_row ["with", "grid", "style"], :style => gridstyle_border
      #sheet.add_row ["no", "border"], :style => [style_test,gridstyle_border]
    #end
  #end
#end
package.serialize 'no_grid_with_borders2.xlsx'

