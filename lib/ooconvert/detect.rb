module OOConvert

  # @return [Symbol] Type of office document this file is recognised as: +:spreadsheet+, +:text+, +:presentation+, +:formula+ or +nil+.
  # @see https://wiki.openoffice.org/wiki/Documentation/OOo3_User_Guides/Getting_Started/File_formats
  def self.type(filename)
    if filename.match /\.(ods|ots|sxc|stc|xls|xlw|xlt|xlsx|xlsm|xltx|xltm|xlsb|wk1|wks|123|sdc|vor|uos|uof|pxl|wb2)$/i
      :spreadsheet
    elsif filename.match /\.(odt|ott|oth|odm|sxw|stw|sxg|doc|dot|docx|docm|dotx|wpd|wps|rtf|sdw|sgl|uot|uof|jtd|jtt|hwp|pdb|psw)$/i
      :text
    elsif filename.match /\.(odp|odg|otp|sxi|sti|ppt|pps|pot|pptx|pptm|potx|potm|sda|sdd|sdo|vor|uop|uof)$/i
      :presentation
    elsif filename.match /\.(odf|sxm|smf|mml)$/i
      :formula
    end
  end

end
