require 'rubygems'
require 'stringio'
require 'spreadsheet'
require 'to_xls/util/hash_merger.rb'

module ToXls

  class Writer
    def initialize(array, options = {})
      @array         = array
      @options       = options
      @cell_format   = create_format :cell_format
      @header_format = create_format :header_format
      @column_format = @options.delete(:column_format) || {}
      @column_width  = @options.delete(:column_width) || {}
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
      if columns.any?
        row_index = 0

        if headers_should_be_included?
          apply_format_to_row(sheet.row(0), @header_format)
          fill_row(sheet.row(0), headers)
          row_index = 1
        end

        @array.each do |model|
          row = sheet.row(row_index)
          apply_format_to_row(row, @cell_format)
          fill_row(row, columns, model)
          row_index += 1
        end

        add_format_to_columns(sheet)
        add_width_to_columns(sheet)
      end
    end

    def columns
      return @columns if @columns
      @columns = @options[:columns]
      raise ArgumentError.new(":columns (#{columns}) must be an array or nil") unless (@columns.nil? || @columns.is_a?(Array))
      @columns ||= can_get_columns_from_first_element? ? get_columns_from_first_element : []
    end

    def can_get_columns_from_first_element?
      @array.first &&
        @array.first.respond_to?(:attributes) &&
        @array.first.attributes.respond_to?(:keys) &&
        @array.first.attributes.keys.is_a?(Array)
    end

    def get_columns_from_first_element
      @array.first.attributes.keys.sort_by {|sym| sym.to_s}.collect.to_a
    end

    def headers
      return @headers if @headers
      @headers = @options[:headers] || columns
      raise ArgumentError, ":headers (#{@headers.inspect}) must be an array" unless @headers.is_a? Array
      @headers
    end

    def headers_should_be_included?
      @options[:headers] != false
    end

  private

    def add_format_to_columns(sheet)
      apply_format_to_all_columns(sheet, column_format_hash.delete(:all))

      column_format_hash.each_pair do |column_name, options|
        column_number = columns.index(column_name)
        apply_format_to_column(sheet.column(column_number), options) if column_number
      end
    end

    def add_width_to_columns(sheet)
      apply_width_to_all_columns(sheet, column_width_hash.delete(:all))

      column_width_hash.each_pair do |column_name, width|
        column_number = columns.index(column_name)
        apply_width_to_column(sheet.column(column_number), width) if column_number
      end
    end

    def apply_format_to_row(row, format)
      row.default_format = format if format
    end

    def apply_format_to_column(column, hash)
      column.default_format = Spreadsheet::Format.new(hash) if hash
    end

    def apply_format_to_all_columns(sheet, value_for_all)
      if value_for_all
        (0...columns.size).each do |column_number|
          apply_format_to_column(sheet.column(column_number), value_for_all)
        end
      end
    end

    def apply_width_to_column(column, width)
      column.width = width if width
    end

    def apply_width_to_all_columns(sheet, value_for_all)
      if value_for_all
        (0...columns.size).each do |column_number|
          apply_width_to_column(sheet.column(column_number), value_for_all)
        end
      end
    end

    def create_format(name)
      Spreadsheet::Format.new @options[name] if @options.has_key? name
    end

    def fill_row(row, column, model=nil)
      case column
      when String, Symbol
        row.push(model ? model.send(column) : column)
      when Hash
        column.each{|key, values| fill_row(row, values, model && model.send(key))}
      when Array
        column.each{|value| fill_row(row, value, model)}
      else
        raise ArgumentError, "column #{column} has an invalid class (#{ column.class })"
      end
    end

    def column_numbers column_names
      if column_names == :all
        (0...columns.size).to_a
      else
        [*column_names].collect{|c| columns.index c }.compact
      end
    end

    def column_format_hash
      @column_format_hash ||= HashMerger.new(@column_format).merge
    end

    def column_width_hash
      @column_width_hash ||= HashMerger.new(@column_width).merge
    end

  end

end
