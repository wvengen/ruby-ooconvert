module OOConvert
  module Encoding
    class << self

      # List of Ruby encodings with Openoffice index and identifier
      # @see https://wiki.openoffice.org/wiki/Documentation/DevGuide/Spreadsheets/Filter_Options
      # @see http://web.archive.org/web/20060507235505/http://equalitylearning.org/~robertz/oo_encoding_indexes.html
      ENCODING_TABLE = {
        'Windows-1252' => [1, 'MS_1252'],
        #Apple Macintosh (Western) 2
        'IBM437' => [3, 'IBM_437'],
        'IBM850' => [4, 'IBM_850'],
        'IBM860' => [5, 'IBM_860'],
        'IBM861' => [6, 'IBM_861'],
        'IBM863' => [7, 'IBM_863'],
        'IBM865' => [8, 'IBM_865'],
        #System default 9
        #Symbol 10
        'US-ASCII' => [11, 'ASCII_US'],
        'ISO-8859-1' => [12, 'ISO_8859_1'],
        'ISO-8859-2' => [13, 'ISO_8859_2'],
        'ISO-8859-3' => [14, 'ISO_8859_3'],
        'ISO-8859-4' => [15, 'ISO_8859_4'],
        'ISO-8859-5' => [16, 'ISO_8859_5'],
        'ISO-8859-6' => [17, 'ISO_8859_6'],
        'ISO-8859-7' => [18, 'ISO_8859_7'],
        'ISO-8859-8' => [19, 'ISO_8859_8'],
        'ISO-8859-9' => [20, 'ISO_8859_9'],
        'ISO-8859-14' => [21, 'ISO_8859_14'],
        'ISO-8859-15' => [22, 'ISO_8859_15'],
        'IBM737' => [23, 'IBM_737'],
        'IBM775' => [24, 'IBM_775'],
        'IBM852' => [25, 'IBM_852'],
        'IBM855' => [26, 'IBM_855'],
        'IBM857' => [27, 'IBM_857'],
        'IBM862' => [28, 'IBM_862'],
        'IBM864' => [29, 'IBM_864'],
        'IBM866' => [30, 'IBM_866'],
        'IBM869' => [31, 'IBM_869'],
        'Windows-874' => [32, 'MS_874'],
        'Windows-1250' => [33, 'MS_1250'],
        'Windows-1251' => [34, 'MS_1251'],
        'Windows-1253' => [35, 'MS_1253'],
        'Windows-1254' => [36, 'MS_1254'],
        'Windows-1255' => [37, 'MS_1255'],
        'Windows-1256' => [38, 'MS_1256'],
        'Windows-1257' => [39, 'MS_1257'],
        'Windows-1258' => [40, 'MS_1258'],
        #Apple Macintosh (Arabic) 41
        'macCentEuro' => [42, 'APPLE_CENTEURO'],
        'macCroatian' => [43, 'APPLE_CROATIAN'],
        'macCyrillic' => [44, 'APPLE_CYRILLIC'],
        #Not supported: Apple Macintosh (Devanagari) 45
        #Not supported: Apple Macintosh (Farsi) 46
        'macGreek' => [47, 'APPLE_GREEK'],
        #Not supported: Apple Macintosh (Gujarati) 48
        #Not supported: Apple Macintosh (Gurmukhi) 49
        #Apple Macintosh (Hebrew)  50
        #Apple Macintosh/Icelandic (Western) 51
        'macRomania' => [52, 'APPLE_ROMANIAN'],
        'macThai' => [53, 'APPLE_THAI'],
        'macTurkish' => [54, 'APPLE_TURKISH'],
        'macUkraine' => [55, 'APPLE_UKRAINIAN'],
        #Apple Macintosh (Chinese Simplified) 56
        #Apple Macintosh (Chinese Traditional)  57
        'MacJapanese' => [58, 'APPLE_JAPANESE'],
        #Apple Macintosh (Korean) 59
        # TODO
        'CP932' => [60, 'MS_932'],
        'CP936' => [61, 'MS_936'],
        'CP949' => [62, 'MS_949'],
        'CP950' => [63, 'MS_950'],
        'Shift_JIS' => [64, 'SHIFT_JIS'],
        'GB2312' => [65, 'GB_2312'],
        'GB12345' => [66, 'GB_12345'],
        'GB2312' => [67, 'GBK'],
        'Big5' => [68, 'BIG5'],
        'EUC-JP' => [69, 'EUC_JP'],
        'EUC-CN' => [70, 'EUC_CN'],
        'EUC-TW' => [71, 'EUC_TW'],
        'ISO-2022-JP' => [72, 'ISO_2022_JP'],
        #ISO-2022-CN (Chinese Simplified) 73
        'KOI8-R' => [74, 'KOI8_R'],
        'UTF-7' => [75, 'UTF7'],
        'UTF-8' => [76, 'UTF8'],
        'ISO-8859-10' => [77, 'ISO_8859_10'],
        'ISO-8859-13' => [78, 'ISO_8859_13'],
        'EUC-KR' => [79, 'EUC_KR'],
        #ISO-2022-KR (Korean) 80
        #JIS 0201 (Japanese) 81
        #JIS 0208 (Japanese) 82
        #JIS 0212 (Japanese) 83
        #Windows-Johab-1361 (Korean) 84
        'GB18030' => [85, 'GB_18030'],
        'Big5-HKSCS' => [86, 'BIG5_HKSCS'],
        'TIS-620' => [87, 'TIS_620'],
        'KOI8-U' => [88, 'KOI8_U'],
        #ISCII Devanagari (Indian) 89
        #Unicode (Java's modified UTF-8) 90
        #Adobe Standard 91
        #Adobe Symbol 92
        #PT 154 (Windows Cyrillic Asian codepage developed in ParaType) 93
        'UCS4-BE' => [65534, 'UCS4'], # BE or LE?? / guessed oo_name
        'UCS2-BE' => [65535, 'UCS2']  # guessed oo_name
      }

      # @param s [Encoding, String] Encoding as used in Ruby
      # @return [Number] Openoffice index for encoding (as used in file filters)
      def index(s)
        ENCODING_TABLE[s.to_s][0]
      end

      # @param s [Encoding, String] Encoding as used in Ruby
      # @return [Number] Openoffice name for encoding (as used in file filters)
      def oo_name(s)
        ENCODING_TABLE[s.to_s][1]
      end

    end
  end
end
