require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe ToXls::Writer do

  it "throws no error without data" do
    lambda { [].to_xls }.should_not raise_error
  end

  describe ":name option" do
    it "defaults to 'Sheet 1' for sheets with no name" do
      make_book([]).worksheets.first.name.should == 'Sheet 1'
    end

    it "uses the :name option" do
      make_book([], :name => 'Empty').worksheets.first.name.should == 'Empty'
    end
  end

  describe ":columns option" do
    it "throws no error without columns" do
      lambda { make_book([1,2,3]) }.should_not raise_error
    end
    it "throws an error if columns isn't an array" do
      lambda { make_book([1,2,3], :columns => :foo) }.should raise_error
    end
    it "uses the attribute keys as columns if it exists" do
      xls = make_book(mock_users)
      check_sheet( xls.worksheets.first,
        [ [:age,  :email,           :name],
          [   20, 'peter@gmail.com', 'Peter'],
          [   25, 'john@gmail.com',  'John'],
          [   27, 'day9@day9tv.com', 'Day9']
        ]
      )
    end
    it "allows re-sorting of the columns by using the :columns option" do
      xls = make_book(mock_users, :columns => [:name, :email, :age])
      check_sheet( xls.worksheets.first,
        [ [:name,   :email,          :age],
          ['Peter', 'peter@gmail.com', 20],
          ['John',  'john@gmail.com',  25],
          ['Day9',  'day9@day9tv.com', 27]
        ]
      )
    end

    it "works properly when you provide it with both data and column names" do
      xls = make_book([1,2,3], :columns => [:to_s])
      check_sheet( xls.worksheets.first, [ [:to_s], ['1'], ['2'], ['3'] ] )
    end

    it "picks data from associations" do
      xls = make_book(mock_users, :columns => [:name, {:company => [:name]}])
      check_sheet( xls.worksheets.first,
        [ [:name,  :name],
          ['Peter', 'Acme'],
          ['John',  'Acme'],
          ['Day9',  'EADS']
        ]
      )
    end
  end

  describe ":headers option" do

    it "uses the headers option if it exists" do
      xls = make_book( mock_users,
        :columns => [:name, :email, :age],
        :headers => ['Nombre', 'Correo', 'Edad']
      )
      check_sheet( xls.worksheets.first,
        [ ['Nombre', 'Correo',      'Edad'],
          ['Peter',  'peter@gmail.com', 20],
          ['John',   'john@gmail.com',  25],
          ['Day9',   'day9@day9tv.com', 27]
        ]
      )
    end

    it "includes no headers if the headers option is false" do
      xls = make_book( mock_users,
        :columns => [:name, :email, :age],
        :headers => false
      )
      check_sheet( xls.worksheets.first,
        [ ['Peter',  'peter@gmail.com', 20],
          ['John',   'john@gmail.com',  25],
          ['Day9',   'day9@day9tv.com', 27]
        ]
      )
    end

    it "picks data from associations" do
      book = make_book( mock_users,
        :columns => [:name, {:company => [:name]}],
        :headers => [:name, :company_name]
      )
      check_sheet( book.worksheets.first,
        [ [:name,  :company_name],
          ['Peter', 'Acme'],
          ['John',  'Acme'],
          ['Day9',  'EADS']
        ]
      )
    end

  end

  describe "#write_book" do
    it "writes a new sheet in a book" do
      book = Spreadsheet::Workbook.new
      ToXls::Writer.new(mock_users).write_book(book)
      check_sheet( book.worksheets.first,
        [ [:age,  :email,           :name],
          [   20, 'peter@gmail.com', 'Peter'],
          [   25, 'john@gmail.com',  'John'],
          [   27, 'day9@day9tv.com', 'Day9']
        ]
      )
    end
  end

  describe "#write_sheet" do
    it "writes a new sheet in a book" do
      book = Spreadsheet::Workbook.new
      sheet = book.create_worksheet
      ToXls::Writer.new(mock_users).write_sheet(sheet)
      check_sheet( sheet,
        [ [:age,  :email,           :name],
          [   20, 'peter@gmail.com', 'Peter'],
          [   25, 'john@gmail.com',  'John'],
          [   27, 'day9@day9tv.com', 'Day9']
        ]
      )
    end
  end

  describe "#write_io" do
    it "writes a new book in a stream" do
      io1 = StringIO.new
      ToXls::Writer.new(mock_users).write_io(io1)
      io2 = StringIO.new
      xls = make_book(mock_users, {})
      xls.write(io2)
      
      io1.string.bytes.to_a.should == io2.string.bytes.to_a
    end
  end

  describe "#write_string" do
    it "writes a new sheet in a string" do
      str = ToXls::Writer.new(mock_users).write_string()
      io = StringIO.new
      xls = make_book(mock_users, {})
      xls.write(io)
      
      str.bytes.to_a.should == io.string.bytes.to_a
    end
  end

  context 'formats' do
    before :each do
      @cell_format   = { :color => :blue }
      @header_format = { :weight => :bold, :color => :red }
      @column_width  = { :age => 4, [:email, :name, :other] => 20 }
      @column_format = { :age => { :number_format => '0.00' }, 
                        [:email, :name, :other] => { :rotation => 90 } }

      @xls = make_book(
        mock_users,
        :columns       => [:name, :age, :email],
        :cell_format   => @cell_format,
        :header_format => @header_format,
        :column_format => @column_format,
        :column_width  => @column_width
      )
    end

    it 'adds header format options to default header format' do
      default_header_format = @xls.worksheets.first.rows[0].default_format
      @header_format.each do |key, value|
        default_header_format.font.send(key).should == value
      end
    end

    it 'adds cell format options to default cell format' do
      default_cell_format = @xls.worksheets.first.rows[1].default_format
      @cell_format.each do |key, value|
        default_cell_format.font.send(key).should == value
      end
    end

    it 'adds column format options to default column format' do
      name_column  = 0
      age_column   = 1
      email_column = 2

      default_column_format = @xls.worksheets.first.columns[age_column].default_format
      default_column_format.number_format.should == '0.00'

      default_column_format = @xls.worksheets.first.columns[email_column].default_format
      default_column_format.rotation.should == 90

      default_column_format = @xls.worksheets.first.columns[name_column].default_format
      default_column_format.rotation.should == 90
    end

    it 'accepts column width option' do
      name_column  = 0
      age_column   = 1
      email_column = 2

      @xls.worksheets.first.columns[age_column].width.should == 4
      @xls.worksheets.first.columns[email_column].width.should == 20
      @xls.worksheets.first.columns[name_column].width.should == 20
    end
  end
end
