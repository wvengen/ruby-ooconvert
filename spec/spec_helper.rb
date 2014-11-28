require 'bundler/setup'
require 'ooconvert'
require 'tempfile'

# run block with temporary file
def with_tempfile
  tempfile = Tempfile.new("ooconvert-test")
  begin
    yield tempfile
  ensure
    tempfile.close
    tempfile.unlink
  end
end
