
###
# Base class describing how a spreadsheet cell should operate
#
class Roo::Base::Cell
  attr_accessor :value
  attr_accessor :type

  def is_header?
  end

  def header
  end
end