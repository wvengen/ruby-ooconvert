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
    # Runs Openoffice and converts the document to the specified format.
    # When no format is specified, it is derived from the destination filename +dst+.
    # If the format is specified, +dst+ may also be a directory, in which case the
    # filename to convert to will be the same as the source document, but with another
    # extension.
    #
    # @param src [String] Filename of source document.
    # @param dst [String] Filename to convert to.
    # @option options [String] :format File format or file extension to convert to.
    # @return [String] Filename of converted document.
    def convert(src, dst=nil, options={})
      dst ||= File.dirname(src)
      options[:format] ||= File.extname(dst)
      convert_to_tmp(src, options) do |tmp|
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
    # @option options [String] :format File format or file extension to convert to (mandatory option).
    def convert_to_tmp(src, options={})
      unless options[:format] && options[:format] != ''
        raise ArgumentError.new('Please specify format')
      end
      src = src.path if src.is_a? File
      Dir.mktmpdir do |tmpdir|
        tmpdst = nil
        tmpsrc = File.join(tmpdir, File.basename(src))
        begin
          FileUtils.copy_file(src, tmpsrc)
          # convert, figure out new filename and call block
          format = Filters.filter_spec(options) or raise ArgumentError.new('Format or extension not recognised')
          convert_to_dir tmpsrc, tmpdir, format
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
    def convert_to_dir(src, dstdir, format)
      # TODO quote
      %x(#{office_bin} --headless --nolockcheck --convert-to '#{format}' '#{src}' --outdir '#{dstdir}' >/dev/null)
    end

  end
end
