require 'tmpdir'
require 'mkmf'

module OOConvert

  class OfficeNotFoundException < Exception
    def message; "openoffice, libreoffice and soffice not found in PATH"; end
  end

  # @return [String] Path to openoffice executable
  # @todo don't use undocumented +find_executable0+ (but it does keep quiet)
  def self.office_bin
    exts = ENV['PATHEXT'] ? ENV['PATHEXT'].split(';') : ['']
    %w(libreoffice openoffice soffice).each do |office|
      exts.each do |ext|
        executable = office + ext
        return executable if find_executable0 executable
      end
    end
    raise OfficeNotFoundException
  end

end
