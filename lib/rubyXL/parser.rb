require 'rubygems'
require 'nokogiri'
require 'zip/zip'
require 'date'
require File.expand_path(File.join(File.dirname(__FILE__),'Hash'))

module RubyXL
  
  # Modified from https://gist.github.com/827475
  #
  # A small DSL for helping parsing documents using Nokogiri::XML::Reader. The
  # XML Reader is a good way to move a cursor through a (large) XML document fast,
  # but is not as cumbersome as writing a full SAX document handler. Read about
  # it here: http://nokogiri.org/Nokogiri/XML/Reader.html
  # 
  # Just pass the reader in this parser and specificy the nodes that you are interested
  # in in a block. You can just parse every node or only look inside certain nodes.
  # 
  # A small example:
  # 
  # Reader.new(filename) do
  #   inside_element 'User' do
  #     for_element 'Name' do puts "Username: #{inner_xml}" end
  #     for_element 'Email' do puts "Email: #{inner_xml}" end
  #     
  #     for_element 'Address' do
  #       puts 'Start of address:'
  #       inside_element do
  #         for_element 'Street' do puts "Street: #{inner_xml}" end
  #         for_element 'Zipcode' do puts "Zipcode: #{inner_xml}" end
  #         for_element 'City' do puts "City: #{inner_xml}" end
  #       end
  #       puts 'End of address'
  #     end
  #   end
  # end
  # 
  # It does NOT fail on missing tags, and does not guarantee order of execution. It parses
  # every tag regardless of nesting. The only way to guarantee scope is by using
  # the `inside_element` method. This limits the parsing to the current or the named tag.
  # If tags are encountered multiple times, their blocks will be called multiple times.
  #
  class Reader
    def initialize(filename, &block)
      @node = Nokogiri::XML::Reader(File.new(filename))
      @node.each do
        self.instance_eval &block
      end
    end
    
    def current_cell
      @current_cell
    end

    def current_cell=(val)
      @current_cell = val
    end

    def name
      @node.name
    end

    def inner_xml
      @node.inner_xml.strip
    end

    def outer_xml
      @node.outer_xml.strip
    end
    
    def value
      value? ? @node.value : nil
    end

    def value?
      @node.value?
    end

    def is_start?
      @node.node_type == Nokogiri::XML::Reader::TYPE_ELEMENT
    end

    def is_end?
      @node.node_type == Nokogiri::XML::Reader::TYPE_END_ELEMENT
    end

    def attribute(attribute)
      @node.attribute(attribute)
    end

    def for_element(name, &block)
      return unless self.name == name and is_start?
      self.instance_eval &block
    end

    def inside_element(name=nil, &block)
      return if @node.self_closing?
      return unless name.nil? or (self.name == name and is_start?)

      name = @node.name
      depth = @node.depth

      @node.each do
        return if self.name == name and is_end? and @node.depth == depth
        self.instance_eval &block
      end
    end
  end

  class Parser
    @parsed_column_hash = {}
    @data_only = false
    @read_only = false
    
    def self.data_only
      @data_only
    end
    
    def self.data_only=(opt)
      @data_only = opt
    end

    def self.read_only
      @read_only
    end
    
    def self.read_only=(opt)
      @read_only = opt
    end
    
    # converts cell string (such as "AA1") to matrix indices
    def self.convert_to_index(cell_string)
      index = Array.new(2)
      index[0] = -1
      index[1] = -1
      if cell_string =~ /^([A-Z]+)(\d+)$/
        one = $1
        row = $2.to_i - 1 #-1 for 0 indexing
        col = 0
        i = 0

        if @parsed_column_hash[one].nil?
          two = one.reverse #because of 26^i calculation
          two.each_byte do |c|
            int_val = c - 64 #converts A to 1
            col += int_val * 26**(i)
            i += 1
          end
          @parsed_column_hash[one] = col
        else
          col = @parsed_column_hash[one]
        end

        col -= 1 #zer0 index
        index[0] = row
        index[1] = col
      end
      return index
    end

    # data_only allows only the sheet data to be parsed, so as to speed up parsing
    #   However, using this option will result in date-formatted cells being interpreted as numbers
    # read_only disables modification or writing of the file, but results in a much
    #   lower memory footprint
    def self.parse(file_path, data_only=false, read_only=false)
      # ensure we are given a xlsx/xlsm file
      if file_path =~ /(.+)\.xls(x|m)/
        dir_path = $1.to_s
      else
        raise 'Not .xlsx or .xlsm excel file'
      end

      self.data_only = data_only
      self.read_only = read_only

      # copy excel file to zip file in same directory
      dir_path = File.join(File.dirname(dir_path), self.make_safe_name(Time.now.to_s))
      zip_path = dir_path + '.zip'
      FileUtils.cp(file_path, zip_path)
      MyZip.new.unzip(zip_path, dir_path)
      File.delete(zip_path)

      # build empty workbook
      wb = Workbook.new([nil], file_path)

      # parse workbook.xml
      self.parse_xml_file(File.join(dir_path, 'xl', 'workbook.xml')) do |workbook_xml|
        # build worksheets
        @num_sheets = Integer(workbook_xml.css('sheets').children.size)
        wb.worksheets = Array.new(@num_sheets)

        # extract everything we need from workbook.xml
        wb.defined_names = workbook_xml.css('definedNames').to_s
        wb.date1904 = workbook_xml.css('workbookPr').attribute('date1904').to_s == '1'
      end

      sheet_names = nil
      self.parse_xml_file(File.join(dir_path, 'docProps', 'app.xml')) do |app_xml|
        sheet_names = app_xml.css('TitlesOfParts vt|vector vt|lpstr').children
        unless self.data_only
          # extract everything we need from app.xml
          wb.company = app_xml.css('Company').children.to_s
          wb.application = app_xml.css('Application').children.to_s
          wb.appversion = app_xml.css('AppVersion').children.to_s
        end
      end

      # parse styles
      unless self.data_only
        self.parse_xml_file(File.join(dir_path, 'xl', 'styles.xml')) do |styles_xml|
          self.fill_styles(wb, Hash.xml_node_to_hash(styles_xml.root))
        end
      end

      # extract everything we need from core.xml
      unless self.data_only
        self.parse_xml_file(File.join(dir_path, 'docProps', 'core.xml')) do |core_xml|
          wb.creator = core_xml.css('dc|creator').children.to_s
          wb.modifier = core_xml.css('cp|last_modified_by').children.to_s
          wb.created_at = core_xml.css('dcterms|created').children.to_s
          wb.modified_at = core_xml.css('dcterms|modified').children.to_s
        end
      end

      # preserve external links
      unless self.data_only
        wb.external_links = self.read_external_files(File.join(dir_path, 'xl', 'externalLinks'))
        wb.drawings = self.read_external_files(File.join(dir_path, 'xl', 'drawings'))
        wb.printer_settings = self.read_external_files(File.join(dir_path, 'xl', 'printerSettings'))
        wb.worksheet_rels = self.read_external_files(File.join(dir_path, 'xl', 'worksheets', '_rels'))
        wb.macros = self.read_external_files(File.join(dir_path, 'xl', 'vbaProject.bin'))
      end

      # extract everything we need from sharedStrings.xml
      wb.shared_strings = {}
      self.parse_shared_strings(wb, File.join(dir_path, 'xl', 'sharedStrings.xml'))

      # parse the worksheets
      for i in 0..@num_sheets-1
        filename = 'sheet' + (i+1).to_s + '.xml'
        wb.worksheets[i] = self.parse_worksheet(wb, File.join(dir_path, 'xl', 'worksheets', filename))
        wb.worksheets[i].sheet_name = sheet_names[i].to_s
      end

      wb.dir_path = dir_path

      wb
    end

    private
    
    def self.parse_shared_strings(wb, filename)
      # Store the whole file for later in case we need to write it out
      if !self.read_only and File.exists?(filename)
        File.open(filename, 'rb') do |f|
          wb.shared_strings_XML = f.read
        end
      end

      i = 0

      # Parse the file
      Reader.new(filename) do
        for_element 'sst' do
          wb.num_strings = Integer(attribute('count'))
          wb.size = Integer(attribute('uniqueCount'))
          
          inside_element do
            for_element 'si' do
              inside_element do
                inside_element 't' do
                  if value?
                    wb.shared_strings[i] = value
                    wb.shared_strings[value] = i unless RubyXL::Parser.read_only
                  end
                end
              end
              i += 1
            end
          end
        end
      end

      # Merge si > r > t into si > t; unnecessary?
      #shared_strings_xml.css('si').each_with_index do |node, i|
      #  unless node.css('r').empty?
      #    text = node.css('r t').children.to_a.join
      #    node.children.remove
      #    node << "<t xml:space=\"preserve\">#{text}</t>"
      #  end
      #end
    end

    # parse worksheet
    def self.parse_worksheet(wb, filename)
      worksheet = Worksheet.new(wb)
      
      Reader.new(filename) do
        inside_element 'worksheet' do
          unless RubyXL::Parser.data_only
            for_element 'sheetViews' do
              h = RubyXL::Parser.xml_to_hash(outer_xml)
              worksheet.sheet_view = h[:sheetView] unless h.nil?
              worksheet.pane = worksheet.sheet_view[:pane] unless worksheet.sheet_view.nil?
            end

            for_element 'cols' do
              h = RubyXL::Parser.xml_to_hash(outer_xml)
              worksheet.cols = h[:col] unless h.nil?
            end

            for_element 'mergeCells' do
              h = RubyXL::Parser.xml_to_hash(outer_xml)
              worksheet.merged_cells = h[:mergeCell] unless h.nil?
            end

            for_element 'dataValidations' do
              h = RubyXL::Parser.xml_to_hash(outer_xml)
              worksheet.validations = h[:dataValidation] unless h.nil?
            end

            for_element 'extLst' do
              worksheet.extLst = RubyXL::Parser.xml_to_hash(outer_xml)
            end

            for_element 'legacyDrawing' do
              worksheet.legacy_drawing = RubyXL::Parser.xml_to_hash(outer_xml)
            end
          end
          
          inside_element 'sheetData' do
            inside_element 'row' do
              if is_start? and !RubyXL::Parser.data_only
                worksheet.row_styles[attribute('r')] = { :style => (attribute('s') || '0')  }

                if !attribute('ht').nil? && !attribute('ht').strip == ""
                  worksheet.change_row_height(Integer(attribute('r'))-1, Float(attribute('ht')))
                end
              end

              for_element 'c' do
                # Get cell index
                current_cell = Parser.convert_to_index(attribute('r'))

                # Add cell if it doesn't exist
                worksheet.sheet_data[current_cell[0]][current_cell[1]] ||= Cell.new(worksheet, current_cell[0], current_cell[1])
                worksheet.sheet_data[current_cell[0]][current_cell[1]].datatype = attribute('t')
                
                inside_element do
                  inside_element 'v' do
                    # Coerce cell value to appropriate type
                    cell_data = nil
                    data_type = worksheet.sheet_data[current_cell[0]][current_cell[1]].datatype
                    if data_type == 's' # shared string
                      cell_data = wb.shared_strings[Integer(value)]
                    elsif data_type == 'str' # raw string
                      cell_data = value
                    elsif data_type == 'e' # error
                      cell_data = value
                    elsif !value.nil? && value != ''
                      worksheet.sheet_data[current_cell[0]][current_cell[1]].datatype = ''
                      if value =~ /\./ #is float
                        cell_data = Float(value)
                      else
                        cell_data = Integer(value)
                      end
                    end
                    
                    # Set cell value
                    worksheet.sheet_data[current_cell[0]][current_cell[1]].value = cell_data
                  end

                  inside_element 'f' do
                    # Set cell formula
                    worksheet.sheet_data[current_cell[0]][current_cell[1]].formula = value
                    
                    # Set cell formula attributes
                    worksheet.sheet_data[current_cell[0]][current_cell[1]].formula_attributes = {
                      't' => attribute('t'),
                      'ref' => attribute('ref'),
                      'si' => attribute('si')
                    }
                  end

                  inside_element 's' do
                    # Set style index
                    unless RubyXL::Parser.data_only
                      unless value.nil? || value == ''
                        worksheet.sheet_data[current_cell[0]][current_cell[1]].style_index = value.to_i
                      end
                    end
                  end
                end
              end
            end
          end
        end
      end

      worksheet
    end
    
    # fill hashes for various styles
    def self.fill_styles(wb,style_hash)
      wb.num_fmts = style_hash[:numFmts]

      ###FONTS###
      wb.fonts = {}
      if style_hash[:fonts][:attributes][:count]==1
        style_hash[:fonts][:font] = [style_hash[:fonts][:font]]
      end

      style_hash[:fonts][:font].each_with_index do |f,i|
        wb.fonts[i.to_s] = {:font=>f,:count=>0}
      end

      ###FILLS###
      wb.fills = {}
      if style_hash[:fills][:attributes][:count]==1
        style_hash[:fills][:fill] = [style_hash[:fills][:fill]]
      end

      style_hash[:fills][:fill].each_with_index do |f,i|
        wb.fills[i.to_s] = {:fill=>f,:count=>0}
      end

      ###BORDERS###
      wb.borders = {}
      if style_hash[:borders][:attributes][:count] == 1
        style_hash[:borders][:border] = [style_hash[:borders][:border]]
      end

      style_hash[:borders][:border].each_with_index do |b,i|
        wb.borders[i.to_s] = {:border=>b, :count=>0}
      end

      wb.cell_style_xfs = style_hash[:cellStyleXfs]
      wb.cell_xfs = style_hash[:cellXfs]
      wb.cell_styles = style_hash[:cellStyles]

      wb.colors = style_hash[:colors]

      #fills out count information for each font, fill, and border
      if wb.cell_xfs[:xf].is_a?(::Hash)
        wb.cell_xfs[:xf] = [wb.cell_xfs[:xf]]
      end
      wb.cell_xfs[:xf].each do |style|
        id = style[:attributes][:fontId].to_s
        unless id.nil?
          wb.fonts[id][:count] += 1
        end

        id = style[:attributes][:fillId].to_s
        unless id.nil?
          wb.fills[id][:count] += 1
        end

        id = style[:attributes][:borderId].to_s
        unless id.nil?
          wb.borders[id][:count] += 1
        end
      end
    end

    def self.parse_options()
      opts = Nokogiri::XML::ParseOptions::DEFAULT_XML
      opts |= Nokogiri::XML::ParseOptions::COMPACT if RubyXL::Parser.read_only
      opts
    end
    
    def self.xml_to_hash(blob)
      Hash.xml_node_to_hash(Nokogiri::XML.parse(blob, nil, nil, self.parse_options).root)
    end

    def self.parse_xml_file(path)
      # Open, parse, and store it
      if File.exist?(path)
        File.open(path, 'rb') do |f|
          yield Nokogiri::XML.parse(f, nil, nil, self.parse_options) if block_given?
        end
      end
    end

    def self.read_external_files(path)
      retval = nil
      
      if File.directory?(path)
        retval = {}
        entries = Dir.new(path).entries.reject { |f| File.directory?(File.join(path, f)) || f == ".DS_Store" }
        entries.each_with_index do |filename, i|
          File.open(File.join(path, filename), 'rb') do |f|
            retval[i+1] = f.read
          end
        end
      elsif File.exists?(path)
        File.open(path, 'rb') do |f|
          retval = f.read
        end
      end
      
      retval
    end

    def self.safe_filename(name, allow_mb_chars=false)
      # "\w" represents [0-9A-Za-z_] plus any multi-byte char
      regexp = allow_mb_chars ? /[^\w]/ : /[^0-9a-zA-Z\_]/
      name.gsub(regexp, "_")
    end

    # Turns the passed in string into something safe for a filename
    def self.make_safe_name(name, allow_mb_chars=false)
      ext = safe_filename(File.extname(name), allow_mb_chars).gsub(/^_/, '.')
      "#{safe_filename(name.gsub(ext, ""), allow_mb_chars)}#{ext}".gsub(/\(/, '_').gsub(/\)/, '_').gsub(/__+/, '_').gsub(/^_/, '').gsub(/_$/, '')
    end

  end
end
