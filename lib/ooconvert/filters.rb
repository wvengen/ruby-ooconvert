require 'csv'

module OOConvert
  # Metadata about Openoffice filters
  #
  # @see https://wiki.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_3_0
  # @see https://wiki.openoffice.org/wiki/Documentation/OOo3_User_Guides/Getting_Started/File_formats
  # @see https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options
  # @see https://github.com/dagwieers/unoconv/blob/master/unoconv
  # @see http://www.artofsolving.com/files/DocumentConverter.py
  module Filters
    class << self

      def filter_spec_out(options)
        type = options[:type]
        format = options[:format]
        if f = find_filters(format, type, 'export').first
          spec_options = spec_options_for(f, options)
          if f[5] =~ /\b#{format}\b/i && !spec_options
            # if we have a file extension match, let OO figure out what's best
            format
          else
            # otherwise use the specified filter with first extension listed
            "#{f[5].split.first rescue nil}:#{f[3]}" + (spec_options ? ":#{spec_options}" : "")
          end
        end
      end


      protected

      # type,nr,localized_name,api_name,unoconv,extensions,flags
      def all_filters
        @all_filters ||= CSV.read(File.join(File.dirname(__FILE__), 'filters.csv')).to_a[1..-1]
      end

      # Find filters by one of the names or extension and type with a given flag.
      def find_filters(s, type, flag)
        type = type.to_s
        flag = flag.upcase
        filters = all_filters.select do |row|
          row[0] == type && (row[2] == s || row[3] == s || row[5] =~ /\b#{s}\b/i) && row[6] =~ /\b#{flag}\b/
        end
        # favour PREFERED, punish NOTINFILEDIALOG and NOTINCHOOSER
        filters.sort_by! do |row|
          (row[6] =~ /\bPREFERED\b/ ? 10 : 0) +
            (row[6] =~ /\bNOTINFILEDIALOG\b/ ? -4 : 0) +
            (row[6] =~ /\bNOTINCHOOSER\b/ ? -4 : 0)
        end
      end

      # Return options specifier string for file format, if sensible
      # @todo full year instead of just two digits
      # @todo also other files than spreadsheets
      def spec_options_for(f, options)
        if f[0] == 'spreadsheet' && [33, 22, 5].include?(f[1].to_i)
          # Lotus, dBase, DIF
          Encoding.index(external_encoding(options)).to_s
        elsif f[0] == 'spreadsheet' && f[1].to_i == 16
          # Text CSV
          col_sep = (options[:col_sep] || ',')
          quote_char = (options[:quote_char] || '"')
          [col_sep.ord, quote_char.ord, Encoding.index(external_encoding(options))].join(',')
        elsif (f[0] == 'text' && f[1].to_i == 26) || (f[0] == 'web' && f[1].to_i == 14)
          # Text (encoded)
          Encoding.oo_name(external_encoding(options))
        end
      end

      # Return desired encoding for the converted file
      def external_encoding(options)
        options[:encoding] || ::Encoding.default_external
      end

    end
  end
end
