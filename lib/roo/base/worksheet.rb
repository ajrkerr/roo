
###
# Base class describing how a worksheet should operate
#
class Roo::Base::Worksheet
  
  attr_accessor :name

  def initialize
    @first_row    = nil
    @last_row     = nil
    @first_column = nil 
    @last_column  = nil
    @cells_read   = false
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

  #TODO These next four methods use a lot of repeatable logic.  Refactor.

  ##
  # Returns the number of the first non-empty row
  def first_row(sheet=nil)
    sheet ||= @default_sheet

    read_cells(sheet)

    # Memoized
    if @first_row[sheet]
      return @first_row[sheet]
    end

    sentinel = 999_999 # more than a spreadsheet can hold
    
    result = sentinel
    
    if @cell[sheet]
      @cell[sheet].each_pair do |key, value|
        y = key.first.to_i # _to_string(key).split(',')
        result = [result, y].min if value
      end
    end
    
    result = nil if result == sentinel
    
    @first_row[sheet] = result
    return result
  end

  # returns the number of the last non-empty row
  def last_row(sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    if @last_row[sheet]
      return @last_row[sheet]
    end
    impossible_value = 0
    result = impossible_value

    @cell[sheet].each_pair do |key,value|
      y = key.first.to_i # _to_string(key).split(',')
      result = [result, y].max if value
    end if @cell[sheet]

    result = nil if result == impossible_value
    @last_row[sheet] = result
    result
  end

  # returns the number of the first non-empty column
  def first_column(sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    if @first_column[sheet]
      return @first_column[sheet]
    end
    impossible_value = 999_999 # more than a spreadsheet can hold
    result = impossible_value
    @cell[sheet].each_pair {|key,value|
      x = key.last.to_i # _to_string(key).split(',')
      result = [result, x].min if value
    } if @cell[sheet]
    result = nil if result == impossible_value
    @first_column[sheet] = result
    result
  end

  # returns the number of the last non-empty column
  def last_column(sheet=nil)
    sheet ||= @default_sheet
    read_cells(sheet)
    if @last_column[sheet]
      return @last_column[sheet]
    end
    impossible_value = 0
    result = impossible_value
    @cell[sheet].each_pair {|key,value|
      x = key.last.to_i # _to_string(key).split(',')
      result = [result, x].max if value
    } if @cell[sheet]
    result = nil if result == impossible_value
    @last_column[sheet] = result
    result
  end

  
end