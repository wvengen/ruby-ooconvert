require 'csv'

module OOConvert
  # Metadata about Openoffice filters
  #
  # @see https://wiki.openoffice.org/wiki/Framework/Article/Filter/FilterList_OOo_3_0
  # @see https://wiki.openoffice.org/wiki/Documentation/OOo3_User_Guides/Getting_Started/File_formats
  # @see https://github.com/dagwieers/unoconv/blob/master/unoconv
  module Filters
    class << self

      def filter_spec(options)
        format = options[:format]
        if f = find_filters(format, 'export').first
          if f[5] =~ /\b#{format}\b/i
            # if we have a file extension match, let OO figure out what's best
            format
          else
            # otherwise use the specified filter with first extension listed
            "#{f[5].split.first rescue nil}:#{f[3]}"
          end
        end
      end


      protected

      # type,nr,localized_name,api_name,unoconv,extensions,flags
      def all_filters
        @all_filters ||= CSV.read(File.join(File.dirname(__FILE__), 'filters.csv')).to_a[1..-1]
      end

      # Find filters by one of the names or extension with a given flag.
      def find_filters(s, flag)
        flag = flag.upcase
        filters = all_filters.select do |row|
          (row[2] == s || row[3] == s || row[5] =~ /\b#{s}\b/i) && row[6] =~ /\b#{flag}\b/
        end
        # @todo sort: favour PREFERED, punish NOTINFILEDIALOG NOTINCHOOSER
        filters
      end

    end
  end
end
