require_relative '../spec_helper'

describe OOConvert::Filters do
  {text: %w(odt ott doc docx), spreadsheet: %w(ods xls xlsx), presentation: %w(ppt pptx)}.each do |type, exts|
    exts.each do |ext|
      it "recognises file-extension '#{ext}'" do
        expect(OOConvert::Filters.filter_spec_out(format: ext, type: type)).to eq ext
      end
    end
  end

  it "recognises txt by api_name" do
    expect(OOConvert::Filters.filter_spec_out(format: 'Text', type: 'text')).to eq 'txt:Text'
  end

  it "knows calc pdf export by api_name" do
    expect(OOConvert::Filters.filter_spec_out(format: 'calc_pdf_Export', type: 'spreadsheet')).to eq 'pdf:calc_pdf_Export'
  end

  it "knows stw by localized_name" do
    expect(OOConvert::Filters.filter_spec_out(format: 'OpenOffice.org 1.0 Text Document Template', type: 'text')).to eq 'stw:writer_StarOffice_XML_Writer_Template'
  end
end
