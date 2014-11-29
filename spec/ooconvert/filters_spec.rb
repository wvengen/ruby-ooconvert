require_relative '../spec_helper'

describe OOConvert::Filters do
  %w(odt doc docx ods xls xlsx odp ppt pptx).each do |ext|
    it "recognises file-extension '#{ext}'" do
      expect(OOConvert::Filters.filter_spec(format: ext)).to eq ext
    end
  end

  it "recognises txt by api_name" do
    expect(OOConvert::Filters.filter_spec(format: 'Text')).to eq 'txt:Text'
  end

  it "knows calc pdf export by api_name" do
    expect(OOConvert::Filters.filter_spec(format: 'calc_pdf_Export')).to eq 'pdf:calc_pdf_Export'
  end

  it "knows stw by localized_name" do
    expect(OOConvert::Filters.filter_spec(format: 'OpenOffice.org 1.0 Text Document Template')).to eq 'stw:writer_StarOffice_XML_Writer_Template'
  end
end
