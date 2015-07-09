require 'rubygems'
require 'stringio'
require 'spreadsheet'
require_relative 'sheet_writer'

module ToXls

  class Writer
    def initialize(array, options = {})
      @array         = array
      @options       = options
    end

    def write_string(string = '')
      io = StringIO.new(string)
      write_io(io)
      io.string
    end

    def write_io(io)
      book = Spreadsheet::Workbook.new
      write_book(book)
      book.write(io)
    end

    def write_book(book)
      sheet = book.create_worksheet
      sheet.name = @options[:name] || 'Sheet 1'
      write_sheet(sheet)
      return book
    end

    def write_sheet(sheet)
      SheetWriter.new(sheet, @array, @options).write
    end

  end

end
