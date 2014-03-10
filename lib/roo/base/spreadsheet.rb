# encoding: utf-8

require 'tmpdir'
require 'stringio'
require 'open-uri'

begin
  require 'zip/zipfilesystem'
  Roo::ZipFile = Zip::ZipFile
rescue LoadError
  # For rubyzip >= 1.0.0
  require 'zip/filesystem'
  Roo::ZipFile = Zip::File
end



# Base class for all other types of spreadsheets
class Roo::Base
  include Enumerable

  TEMP_PREFIX = "oo_"

  attr_reader :default_sheet, :headers

  # sets the line with attribute names (default: 1)
  attr_accessor :header_line


public

  def initialize(filename, options={}, file_warning=:error, tmpdir=nil)
    @filename = filename
    @options  = options

    # Format for cells is @cell[sheet][[row,col]] = value
    @cell        = {}
    @cell_type   = {}
    @cells_read  = {}

    @first_row    = {}
    @last_row     = {}
    @first_column = {}
    @last_column  = {}

    @header_line   = 1
    @default_sheet = self.sheets.first
  end

  # sets the working sheet in the document
  # 'sheet' can be a number (1 = first sheet) or the name of a sheet.
  def default_sheet=(sheet)
    sheet = get_sheet(sheet)
    @default_sheet = sheet

    @first_row[sheet]    = nil
    @last_row[sheet]     = nil
    @first_column[sheet] = nil 
    @last_column[sheet]  = nil
    @cells_read[sheet]   = false
  end

  ##
  # Returns the sheet
  def get_sheet sheet
    case sheet
    when nil
      return default_sheet

    when Fixnum
      return self.sheets[sheet-1] or raise RangeError, "Sheet index #{sheet} not found"

    when String
      return sheets.include?(sheet) or raise RangeError, "sheet '#{sheet}' not found"

    when Roo::Base::Worksheet
      if sheets.include?(sheet)
        return sheet 
      else
        raise ArgumentError, "Sheet #{sheet.name} is not a valid worksheet for this spreadsheet"
      end

    else
      raise TypeError, "Not a valid sheet type: #{sheet.inspect}"
    end

  end

  ##
  # Returns the first non-empty column from the sheet as a letter
  def first_column_as_letter(sheet=nil)
    Roo::Base.number_to_letter(first_column(sheet))
  end

  ##
  # Returns the last non-empty column from the sheet as a letter
  def last_column_as_letter(sheet=nil)
    Roo::Base.number_to_letter(last_column(sheet))
  end

  ##
  # Returns the number of the first non-empty row
  def first_row(sheet=nil)
    get_sheet(sheet).first_row
  end

  # returns the number of the last non-empty row
  def last_row(sheet=nil)
    get_sheet(sheet).last_row
  end

  # returns the number of the first non-empty column
  def first_column(sheet=nil)
    get_sheet(sheet).first_column
  end

  # returns the number of the last non-empty column
  def last_column(sheet=nil)
    get_sheet(sheet).last_column
  end

  ##
  # find a row either by row number or a condition
  # Caution: this works only within the default sheet -> set default_sheet before you call this method
  # (experimental. see examples in the test_roo.rb file)
  def find(*args) # :nodoc
    get_sheet.find(*args)
  end

  ##
  # Returns all values in this row as an array
  # Rows are numbered starting with 1, like most spreadsheet software
  def row(row_number,sheet=nil)
    get_sheet(sheet).row(row_number)
  end

  ##
  # returns all values in this column as an array
  # column numbers are 1,2,3,... like in the spreadsheet
  def column(col_number,sheet=nil)
    get_sheet(sheet).column(col_number)
  end

  ##
  # set a cell to a certain value
  # (this will not be saved back to the spreadsheet file!)
  def set(row,col,value,sheet=nil) #:nodoc:
    get_sheet(sheet).set(row, col, value)
  end

  ##
  # Reloads the spreadsheet document
  def reload
    ds = @default_sheet
    initialize(@filename)
    self.default_sheet = ds
  end

  ##
  # True if a cell is empty
  def empty?(row, col, sheet=nil)
    get_sheet(sheet).empty?(row, col)
  end

### Output Functions ###
  # returns a rectangular area (default: all cells) as yaml-output
  # you can add additional attributes with the prefix parameter like:
  # oo.to_yaml({"file"=>"flightdata_2007-06-26", "sheet" => "1"})
  def to_yaml(prefix={}, from_row=nil, from_column=nil, to_row=nil, to_column=nil,sheet=nil)
    get_sheet(sheet).to_yaml
    sheet ||= @default_sheet
    result = "--- \n"
    return '' unless first_row # empty result if there is no first_row in a sheet

    (from_row||first_row(sheet)).upto(to_row||last_row(sheet)) do |row|
      (from_column||first_column(sheet)).upto(to_column||last_column(sheet)) do |col|
        unless empty?(row,col,sheet)
          result << "cell_#{row}_#{col}: \n"
          prefix.each {|k,v|
            result << "  #{k}: #{v} \n"
          }
          result << "  row: #{row} \n"
          result << "  col: #{col} \n"
          result << "  celltype: #{self.celltype(row,col,sheet)} \n"
          if self.celltype(row,col,sheet) == :time
            result << "  value: #{Roo::Base.integer_to_timestring( self.cell(row,col,sheet))} \n"
          else
            result << "  value: #{self.cell(row,col,sheet)} \n"
          end
        end
      end
    end
    result
  end

  # write the current spreadsheet to stdout or into a file
  def to_csv(filename=nil,sheet=nil,separator=',')
    sheet ||= @default_sheet
    if filename
      File.open(filename,"w") do |file|
        write_csv_content(file,sheet,separator)
      end
      return true
    else
      sio = StringIO.new
      write_csv_content(sio,sheet,separator)
      sio.rewind
      return sio.read
    end
  end

  # returns a matrix object from the whole sheet or a rectangular area of a sheet
  def to_matrix(from_row=nil, from_column=nil, to_row=nil, to_column=nil,sheet=nil)
    require 'matrix'

    sheet ||= @default_sheet
    return Matrix.empty unless first_row

    Matrix.rows((from_row||first_row(sheet)).upto(to_row||last_row(sheet)).map do |row|
      (from_column||first_column(sheet)).upto(to_column||last_column(sheet)).map do |col|
        cell(row,col,sheet)
      end
    end)
  end

  # returns an XML representation of all sheets of a spreadsheet file
  def to_xml
    Nokogiri::XML::Builder.new do |xml|
      xml.spreadsheet do
        self.sheets.each do |sheet|
          self.default_sheet = sheet
          xml.sheet(:name => sheet) do |x|
            if first_row and last_row and first_column and last_column
              # sonst gibt es Fehler bei leeren Blaettern
              first_row.upto(last_row) do |row|
                first_column.upto(last_column) do |col|
                  unless empty?(row,col)
                    x.cell(cell(row,col),
                      :row =>row,
                      :column => col,
                      :type => celltype(row,col))
                  end
                end
              end
            end
          end
        end
      end
    end.to_xml
  end

  ##
  # returns information of the spreadsheet document and all sheets within
  # this document.
  def info
    without_changing_default_sheet do
      result = "File: #{File.basename(@filename)}\n"+
        "Number of sheets: #{sheets.size}\n"+
        "Sheets: #{sheets.join(', ')}\n"
      n = 1
      sheets.each {|sheet|
        self.default_sheet = sheet
        result << "Sheet " + n.to_s + ":\n"
        unless first_row
          result << "  - empty -"
        else
          result << "  First row: #{first_row}\n"
          result << "  Last row: #{last_row}\n"
          result << "  First column: #{Roo::Base.number_to_letter(first_column)}\n"
          result << "  Last column: #{Roo::Base.number_to_letter(last_column)}"
        end
        result << "\n" if sheet != sheets.last
        n += 1
      }
      result
    end
  end


  ##
  # When a method like spreadsheet.a42 is called
  # Convert it to a call of spreadsheet.cell('a',42)
  def method_missing(method, *args)
    # #aa42 => #cell('aa',42)
    # #aa42('Sheet1')  => #cell('aa',42,'Sheet1')
    if method =~ /^([a-z]+)(\d)$/
      col = Roo::Base.letter_to_number($1)
      row = $2.to_i

      if args.empty?
        cell(row,col)
      else
        cell(row,col,args.first)
      end
    else
      super
    end
  end

  def respond_to?(method, include_all= false)
    # #aa42 => #cell('aa',42)
    # #aa42('Sheet1')  => #cell('aa',42,'Sheet1')
    if method =~ /^([a-z]+)(\d)$/
      true
    else
      super
    end
  end

  ##
  # access different worksheets by calling spreadsheet.sheet(1)
  # or spreadsheet.sheet('SHEETNAME')
  def sheet(index, name=false)
    @default_sheet = String === index ? index : self.sheets[index]
    name ? [@default_sheet,self] : self
  end

  ##
  # Iterate through all worksheets of a document
  def each_with_pagename
    self.sheets.each do |s|
      yield sheet(s, true)
    end
  end

  # by passing in headers as options, this method returns
  # specific columns from your header assignment
  # for example:
  # xls.sheet('New Prices').parse(:upc => 'UPC', :price => 'Price') would return:
  # [{:upc => 123456789012, :price => 35.42},..]

  # the queries are matched with regex, so regex options can be passed in
  # such as :price => '^(Cost|Price)'
  # case insensitive by default


  # by using the :header_search option, you can query for headers
  # and return a hash of every row with the keys set to the header result
  # for example:
  # xls.sheet('New Prices').parse(:header_search => ['UPC*SKU','^Price*\sCost\s'])

  # that example searches for a column titled either UPC or SKU and another
  # column titled either Price or Cost (regex characters allowed)
  # * is the wildcard character

  # you can also pass in a :clean => true option to strip the sheet of
  # odd unicode characters and white spaces around columns

  def each(options={})
    if options.empty?
      1.upto(last_row) do |line|
        yield row(line)
      end
    else
      if options[:clean]
        options.delete(:clean)
        @cleaned ||= {}
        @cleaned[@default_sheet] || clean_sheet(@default_sheet)
      end

      if options[:header_search]
        @headers = nil
        @header_line = row_with(options[:header_search])
      elsif [:first_row,true].include?(options[:headers])
        @headers = []
        row(first_row).each_with_index {|x,i| @headers << [x,i + 1]}
      else
        set_headers(options)
      end

      headers = @headers ||
        Hash[(first_column..last_column).map do |col|
          [cell(@header_line,col), col]
        end]

      @header_line.upto(last_row) do |line|
        yield(Hash[headers.map {|k,v| [k,cell(line,v)]}])
      end
    end
  end

  def parse(options={})
    ary = []
    if block_given?
      each(options) {|row| ary << yield(row)}
    else
      each(options) {|row| ary << row}
    end
    ary
  end

  def row_with(query,return_headers=false)
    query.map! {|x| Array(x.split('*'))}
    line_no = 0
    each do |row|
      line_no += 1
      # makes sure headers is the first part of wildcard search for priority
      # ex. if UPC and SKU exist for UPC*SKU search, UPC takes the cake
      headers = query.map do |q|
        q.map {|i| row.grep(/#{i}/i)[0]}.compact[0]
      end.compact

      if headers.length == query.length
        @header_line = line_no
        return return_headers ? headers : line_no
      elsif line_no > 100
        raise "Couldn't find header row."
      end
    end
  end




protected
  ##
  # 
  def self.split_coordinate(str)
    letter,number = Roo::Base.split_coord(str)
    x = letter_to_number(letter)
    y = number
    return y, x
  end

  ##
  # Attempts to split the row and column from a string
  # eg. "A2" => ["A", 2]
  def self.split_coord(s)
    if s =~ /([a-zA-Z]+)([0-9]+)/
      letter = $1
      number = $2.to_i
    else
      raise ArgumentError
    end
    return letter, number
  end

  ##
  # Loads XML from a path
  def load_xml(path)
    File.open(path) do |file|
      Nokogiri::XML(file)
    end
  end

  # TODO Refactor this file type check
  def file_type_check(filename, ext, name, warning_level, packed=nil)
    new_expression = {
      '.ods'  => 'Roo::OpenOffice.new',
      '.xls'  => 'Roo::Excel.new',
      '.xlsx' => 'Roo::Excelx.new',
      '.xlsm' => 'Roo::Excelx.new',
      '.csv'  => 'Roo::CSV.new',
      '.xml'  => 'Roo::Excel2003XML.new',
    }

    if packed == :zip
      # lalala.ods.zip => lalala.ods
      # Here is NOT made ​​unzip, but only the name of the file 
      # Tested, if it is a compressed file.
      filename = File.basename(filename,File.extname(filename))
    end

    case ext
    when '.ods', '.xls', '.xlsx', '.csv', '.xml', '.xlsm'
      correct_class = "use #{new_expression[ext]} to handle #{ext} spreadsheet files. This has #{File.extname(filename).downcase}"
    else
      raise "unknown file type: #{ext}"
    end

    if uri?(filename) && qs_begin = filename.rindex('?')
      filename = filename[0..qs_begin-1]
    end

    extension = File.extname(filename).downcase

    ## TODO: Rewire this check to not have a hardcoded extension check
    if extension != ext and (extension != '.xlsm' and ext != '.xlsx')
      case warning_level
      when :error
        warn correct_class
        raise TypeError, "#{filename} is not #{name} file"
      when :warning
        warn "are you sure, this is #{name} spreadsheet file?"
        warn correct_class
      when :ignore
        # ignore
      else
        raise "#{warning_level} illegal state of file_warning"
      end
    end
  end

  ##
  # Converts a key of the form "12.45" (= row, column) in 
  # An array with numeric values ​​([12,45])
  # This method is a temp. Solution in order to explore whether the 
  # Access with numeric keys is faster.
  def key_to_num(str)
    r,c = str.split(',')
    [r.to_i,c.to_i]
  end

  # see: key_to_num
  def key_to_string(arr)
    "#{arr[0]},#{arr[1]}"
  end

private

  def find_by_row(args)
    rownum = args[0]
    current_row = rownum
    current_row += header_line - 1 if @header_line

    self.row(current_row).size.times.map do |j|
      cell(current_row, j + 1)
    end
  end

  def find_by_conditions(options)
    rows = first_row.upto(last_row)
    result_array = options[:array]
    header_for = Hash[1.upto(last_column).map do |col|
      [col, cell(@header_line,col)]
    end]

    # are all conditions met?
    if (conditions = options[:conditions]) && !conditions.empty?
      column_with = header_for.invert
      rows = rows.select do |i|
        conditions.all? { |key,val| cell(i,column_with[key]) == val }
      end
    end

    rows.map do |i|
      if result_array
        self.row(i)
      else
        Hash[1.upto(self.row(i).size).map do |j|
          [header_for.fetch(j), cell(i,j)]
        end]
      end
    end
  end


  ##
  # Perform an operation without changing the default spreadsheet
  def without_changing_default_sheet
    original_default_sheet = default_sheet
    yield
  ensure
    self.default_sheet = original_default_sheet
  end

  ##
  # Creates a temproary directory for us to work in
  def make_tmpdir(tmp_folder = nil)
    Dir.mktmpdir(TEMP_PREFIX, tmp_folder || ENV['ROO_TMP']) do |tmpdir|
      yield tmpdir
    end
  end

  def clean_sheet(sheet)
    read_cells(sheet)

    @cell[sheet].each_pair do |coord,value|
      if String === value
        @cell[sheet][coord] = sanitize_value(value)
      end
    end

    @cleaned[sheet] = true
  end

  ##
  # What does this do?
  def sanitize_value(v)
    v.strip.unpack('U*').select {|b| b < 127}.pack('U*')
  end

  def set_headers(hash={})
    # try to find header row with all values or give an error
    # then create new hash by indexing strings and keeping integers for header array
    @headers = row_with(hash.values,true)
    @headers = Hash[hash.keys.zip(@headers.map {|x| header_index(x)})]
  end

  def header_index(query)
    row(@header_line).index(query) + first_column
  end

  def set_value(row,col,value,sheet=nil)
    sheet ||= @default_sheet
    @cell[sheet][[row,col]] = value
  end

  def set_type(row,col,type,sheet=nil)
    sheet ||= @default_sheet
    @cell_type[sheet][[row,col]] = type
  end

  ##
  # converts cell coordinate to numeric values of row,col
  def normalize(row,col)
    if row.class == String
      if col.class == Fixnum
        # ('A',1):
        # ('B', 5) -> (5, 2)
        row, col = col, row
      else
        raise ArgumentError
      end
    end

    if col.class == String
      col = Roo::Base.letter_to_number(col)
    end

    return row, col
  end

  ##
  # TODO: More robust check
  def uri?(filename)
    filename.start_with?("http://", "https://")
  end

  ##
  # Downloads a file from a URL for us to use
  def download_uri(uri)
    tempfile = Tempfile.new(File.basename(uri))

    begin
      open(uri, "User-Agent" => "Ruby/#{RUBY_VERSION}") { |net|
        tempfile.write(net.read)
      }
    rescue OpenURI::HTTPError
      raise "could not open #{uri}"
    end

    tempfile.path
  end

  ##
  # TODO Provide option for workign in memory
  def open_from_stream(stream, tmpdir)
    tempfile = Tempfile.new(File.basename(uri))

    tempfile.write(stream[7..-1])
    
    tempfile.path
  end

  LETTERS = %w{A B C D E F G H I J K L M N O P Q R S T U V W X Y Z}

  ##
  # Convert a number to letters as per the standard spreadsheet columnation 
  # eg. 27 => 'AA' (1 => 'A', 2 => 'B', ...)
  def self.number_to_letter(number)
    result = ""
    number = number.to_i
    
    while number > 0 do
      modulo = (number - 1) % 26
      number = (number - modulo) / 26

      result += LETTERS[modulo]
    end

    return result.reverse
  end

  ##
  # Convert a letters to numbers as per the standard spreadsheet columnation 
  # eg. 'AA' => 27 (1 => 'A', 2 => 'B', ...)
  def self.letter_to_number(letters)
    result = 0

    letters.each_char do |character|
      num = LETTERS.index(character.upcase)
      raise ArgumentError, "invalid column character '#{character}'" if num == nil
      result  = result * 26 + num + 1
    end

    return result
  end

  ##
  # Converts an integer value to a time string like '02:05:06'
  def self.integer_to_timestring(content)
    h = (content/3600.0).floor
    content = content - h*3600
    m = (content/60.0).floor
    content = content - m*60
    s = content
    sprintf("%02d:%02d:%02d", h, m, s)
  end

  def unzip(filename, tmpdir)
    Roo::ZipFile.open(filename) do |zip|
      process_zipfile_packed(zip, tmpdir)
    end
  end

  # check if default_sheet was set and exists in sheets-array
  def validate_sheet!(sheet)
    warn "Depricated ValidateSheet!"
  end

  def process_zipfile_packed(zip, tmpdir, path='')
    if zip.file.file? path
      # extract and return filename
      File.open(File.join(tmpdir, path),"wb") do |file|
        file.write(zip.read(path))
      end
      File.join(tmpdir, path)
    else
      ret=nil
      path += '/' unless path.empty?
      zip.dir.foreach(path) do |filename|
        ret = process_zipfile_packed(zip, tmpdir, path + filename)
      end
      ret
    end
  end

  # Write all cells to the csv file. File can be a filename or nil. If the this
  # parameter is nil the output goes to STDOUT
  def write_csv_content(file=nil,sheet=nil,separator=',')
    file ||= STDOUT
    if first_row(sheet) # sheet is not empty
      1.upto(last_row(sheet)) do |row|
        1.upto(last_column(sheet)) do |col|
          file.print(separator) if col > 1
          file.print cell_to_csv(row,col,sheet)
        end
        file.print("\n")
      end # sheet not empty
    end
  end

  # The content of a cell in the csv output
  def cell_to_csv(row, col, sheet)
    if empty?(row,col,sheet)
      ''
    else
      onecell = cell(row,col,sheet)

      case celltype(row,col,sheet)
      when :string
        unless onecell.empty?
          %{"#{onecell.gsub(/"/,'""')}"}
        end
      when :boolean
        %{"#{onecell.gsub(/"/,'""').downcase}"}
      when :float, :percentage
        if onecell == onecell.to_i
          onecell.to_i.to_s
        else
          onecell.to_s
        end
      when :formula
        case onecell
        when String
          unless onecell.empty?
            %{"#{onecell.gsub(/"/,'""')}"}
          end
        when Float
          if onecell == onecell.to_i
            onecell.to_i.to_s
          else
            onecell.to_s
          end
        when DateTime
          onecell.to_s
        else
          raise "unhandled onecell-class #{onecell.class}"
        end
      when :date, :datetime
        onecell.to_s
      when :time
        Roo::Base.integer_to_timestring(onecell)
      when :link
          %{"#{onecell.url.gsub(/"/,'""')}"}
      else
        raise "unhandled celltype #{celltype(row,col,sheet)}"
      end || ""
    end
  end
end
