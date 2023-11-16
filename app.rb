require "google_drive"

spreadsheet_key = "13AtUm-Dl-ccbvwjEtHmmGl5ULV4LtmqzV7INcgPvm7c";

class String
	def is_integer?
	  self.to_i.to_s == self
	end
end

class Shreadsheet

	# google drive session thing
	@@session = GoogleDrive::Session.from_config("config.json")
	
	# spreadsheet document
	@document
	# spreadsheet active/current sheet
	@worksheet

	# table start row index relative to workseet (header top left position)
	@start_row
	# table start column index relative to workseet (header top left position)
	@start_column
	
	# number of rows in table (includes header)
	@rows
	# number of columns in table
	@columns
	
	# table header names
	@header
	# table values
	@data

	attr_reader :worksheet

	def initialize(spreadsheet_key)
		@document = @@session.spreadsheet_by_key(spreadsheet_key)
		self.worksheet = 0
	end

	def print
		p "dokument | broj kolona: #{worksheet.num_cols}"
		p "dokument | broj redova: #{worksheet.num_rows}"
		p "tabela | broj kolona: #{@columns}"
		p "tabela | broj redova: #{@rows}"
		p "tablea | prva kolona: #{@start_column}"
		p "tabela | prvi red: #{@start_row}"
	end

	def worksheet=(index)
		@worksheet = @document.worksheets[index]
		
		if (@worksheet.nil?)
			p "Worksheet with index #{index} does not exist!"
			return
		end

		self.find_table()
		self.read_data()
	end

	def find_table
		last_row    = @worksheet.num_rows
		last_column = @worksheet.num_cols

		# find number of rows and start row index
		@rows = 1
		while worksheet[last_row - @rows, last_column].empty? ||
			  worksheet[last_row - @rows, last_column].is_integer?
			@rows += 1
		end
		# include header
		@rows += 1
		@start_row = last_row - @rows + 1

		# find number of columns and start column index
		@columns = 1
		while not worksheet[@start_row, last_column - @columns].empty?
			@columns += 1
		end
		@start_column = last_column - @columns + 1
	end

	def read_data
		# read table header data
		@header = []

		column = @start_column
		while not worksheet[@start_row, column].empty?
			@header << worksheet[@start_row, column]
			column += 1;
		end

		# read table values
		@data = Array.new(@rows * @columns - @columns)
		(@start_row + 1..@worksheet.num_rows).each do |row|
			(@start_column..@worksheet.num_cols).each do |column|
				@data[(row - @start_row - 1) * @columns + column - @start_column] =
					worksheet[row, column].to_i if not worksheet[row, column].empty?
			end
		end
		p @data
	end

	def row(row)
		if row >= @rows
			return []
		end

		array = Array.new(@columns)

		(0..@columns - 1).each do |index|
			array[index] = @data[row * @columns + index]
		end

		return array
	end

	private :find_table, :read_data

end

api = Shreadsheet.new(spreadsheet_key)
api.print

api.worksheet = 1
