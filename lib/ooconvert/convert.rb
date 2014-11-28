require 'fileutils'
require 'tmpdir'

module OOConvert

  class Exception < ::Exception; end

  class ConversionFailedException < Exception
    def message; "Conversion to CSV failed"; end
  end

  class << self

    # Convert document.
    #
    # @param src [String] Filename of source document.
    # @param format [String] Format to convert to (as accepted by Openoffice's +convert-to+ argument)
    # @param dst [String] Filename or directory to convert to, or +nil+ to use directory of source file.
    # @return [String] Filename of converted document.
    def convert(src, format, dst=nil)
      dst ||= File.dirname(src)
      convert_to_tmp(src, format) do |tmp|
        if File.directory?(dst)
          dst = File.join(dst, File.basename(tmp))
        end
        FileUtils.move tmp, dst
      end
      dst
    end

    # Convert document to temporary file.
    #
    # Expects a block, which is called with the temporary filename of
    # the converted file as argument. This file is deleted after the
    # block ends, so if you need to keep it, make sure to make a copy.
    #
    # @param src [String, File] Filename to convert
    # @param format [String] Format to convert to (as accepted by Openoffice's +convert-to+ argument)
    def convert_to_tmp(src, format)
      src = src.path if src.is_a? File
      Dir.mktmpdir do |tmpdir|
        tmpdst = nil
        tmpsrc = File.join(tmpdir, File.basename(src))
        begin
          FileUtils.copy_file(src, tmpsrc)
          # convert, figure out new filename and call block
          convert_to_dir tmpsrc, format, tmpdir
          tmpdst = Dir.entries(tmpdir).reject{|f| ['.', '..', File.basename(tmpsrc)].include? f }.first
          tmpdst or raise ConversionFailedException
          tmpdst = File.join(tmpdir, tmpdst)
          yield tmpdst
        ensure
          FileUtils.rm_f tmpsrc
          FileUtils.rm_f tmpdst if tmpdst
        end
      end
      nil
    end


    protected

    # Convert to a specified directory
    def convert_to_dir(src, format, dstdir)
      # TODO quote
      %x(#{office_bin} --headless --nolockcheck --convert-to '#{format}' '#{src}' --outdir '#{dstdir}' >/dev/null)
    end

  end
end
