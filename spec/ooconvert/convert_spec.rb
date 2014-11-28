# encoding:utf-8
require_relative '../spec_helper'
require 'csv'

describe OOConvert do
  describe 'convert' do

    Dir.glob('spec/files/test_01*') do |file|
      it "'#{file}' to csv file" do
        with_tempfile do |tempfile|
          expect(OOConvert.convert(file, 'csv', tempfile.path)).to eq tempfile.path
          expect(CSV.read(tempfile.path, encoding: 'ISO-8859-1:UTF-8').to_a).to eq [
            ["Date", "Name", "Code", "Attending"],
            ["25/10/2010", "Important event", "IIM", "12"],
            ["30/11/2010", "Other event", "AMD", "5"],
            ["01/01/2011", "Happy néw yeâr", "AÇ", "101"]
          ]
        end
      end
    end

    Dir.glob('spec/files/test_02*') do |file|
      it "'#{file}' to text file" do
        with_tempfile do |tempfile|
          expect(OOConvert.convert(file, 'txt:Text', tempfile.path)).to eq tempfile.path
          expect(File.read(tempfile.path, encoding: 'BOM|UTF-8')).to \
            eq "Test document\nThis is a test δοκυμεντ as ¥ see. Bye!\n"
        end
      end
    end

  end
end
