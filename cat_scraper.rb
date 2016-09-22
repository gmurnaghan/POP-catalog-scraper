
class CatScraper
		A_to_I  = { 'a' => 	1,
								'b' => 	2,
								'c' => 	3,
								'd' => 	4,
								'e' => 	5,
								'f' => 	6,
								'g' => 	7,
								'h' => 	8,
								'i' => 	9,
								'j' => 10,
								'k' => 11,
								'l' => 12,
								'm' => 13,
								'n' => 14,
								'o' => 15,
								'p' => 16,
								'q' => 17,
								'r' => 18,
								's' => 19,
								't' => 20,
								'u' => 21,
								'v' => 22,
								'w' => 23,
								'x' => 24,
								'y' => 25,
								'z' => 26 }

	def alph_to_int alph
		int = 0
			(1..alph.length).each do |place|
				a = alph[0 - place]
				int += (A_to_I[a] * (26 ** (place - 1)))
			end
		int
	end

	def int_to_alph i
		alph = 'A'
		(1...i).each do
			 alph = alph.succ
		end
		alph
	end

	def get_spreadsheet
		file_path = ''
		loop do
			puts 'Please enter the location of the Excel file'
			file_path = gets.chomp
			if Pathname.new(file_path).exist?
				break
			else
				puts 'Invalid file location'
			end
		end
		file_path
	end

	def get_columns
	data_range = ''
	loop do
			puts 'Column(s)?'
			cols = gets.chomp.downcase
			if cols  =~ (/^[a-z]+-[a-z]+$/) || cols =~ (/^[a-z]+$/)
				if cols.include? '-'
					first_col = (cols.scan (/(.*)-/))[0][0]
					last_col  = (cols.scan (/-(.*)/))[0][0]
				else
					first_col = cols
					last_col  = cols
				end
				column_nums = [alph_to_int(first_col) , alph_to_int(last_col)]
				if column_nums[0] <= column_nums[1]
					data_range = (column_nums[0] .. column_nums[1])
					break
				else
				puts 'Invalid range (first heading larger than second heading)'
				end
			else
				puts 'Please enter two column headings seperated by a dash (e.g. \'A-ZZ\'), or  a single heading'
			end
		end
		data_range
	end

	def check_cells sheet, column, heading
		filled_cells = ''
		(5..18).each do |row|
			unless row == 10 or sheet.cell(row, column).nil?
				filled_cells	<< "#{row}, "
			end
		end
		unless filled_cells.empty?
			puts "CAUTION, there is already data entered in these rows: #{filled_cells.chop.chop}"
			puts "Skip this column? (Y/N)"
			choice = ''
			loop do
				input = gets.chomp.downcase
				if input == 'y' or input == 'n'
					choice = input
					break
				else
					puts "enter 'Y' or 'N'"
				end
			end
			if choice == 'y'
				true
			else
				print "[#{heading}]: "
				false
			end
		else
			false
		end
	end

	CULTURE_CODES = {'AC'   =>	'American',
									 'BelC' => 'Belgian',
									 'CzC'  => 'Czech',
									 'EC'   => 'English',
									 'FC'   => 'French',
									 'GC'   => 'German',
									 'GrC'  => 'Greek',
									 'HunC' => 'Hungarian',
									 'IC'   => 'Italian',
									 'LatC' => 'Latin',
									 'NC'   => 'Netherlands',
									 'PC'   => 'Portuguese',
									 'PolC' => 'Polish',
									 'RC'   =>	'Russian',
									 'SC'   => 'Spanish',
									 'ScdC' => 'Danish',
									 'ScnC' => 'Norwegian',
									 'ScsC' => 'Swedish',
									 'SwC'  => 'Swiss',
									 'YugC' => 'Yugoslavian'}

	def get_fields call_num, url, record_section, notifications
		# ---------------------------------------------
		current_repository = 'Penn Libraries'
		# ----------------------------------------------
		current_collection = 'COULD NOT FIND'
		if call_num.include? 	 'Inc'
			 current_collection = 'Incunables Collection'
		else
			CULTURE_CODES.each do |pairing|
				if call_num.include? pairing[0]
					current_collection = "#{pairing[1]} Culture Class Collection"
					break
				end
			end
		end
		# ----------------------------------------------
		current_location = 'Philadelphia'
		# ----------------------------------------------
		author = ''
		author_rec = record_section.scan(/>Author\/Creator.*?cet"> (.*?)<\/a>/m)[0]
								 # Nokogiri
		unless author_rec.nil?
			author = author_rec[0]
		end
		if author[-1] == ','
			notifications << 'author credit may not be author'
		end
		# ----------------------------------------------
		title = 'NOT FOUND'
		title_rec = record_section.scan(/<h1 class="recordtitle">(.*?)<\/h1>/m)[0]
								# Nokogiri
		unless title_rec.nil?
			title = title_rec[0].gsub (/ *\n */) , (' ')
		end
		# ----------------------------------------------
		place_of_pub = ''
		pub_rec = record_section.scan(/>Place of c.*?dard">(.*?)<\/a>/m)[0]
							# Nokogiri
		unless pub_rec.nil?
			place_of_pub = pub_rec[0]
		end
		# ----------------------------------------------
		date_narrative = ''
		date_standard  = 'NOT FOUND'
		date_rec = record_section.scan(/>Chrono.*?ject"> (.*?)<\/a>/m)[0]
							# Nokogiri
		unless date_rec.nil?
			date = date_rec[0].chop
			date_standard = date.match(/[0-9][0-9][0-9\?][0-9\?]/).to_s.gsub('?' , '0')
			if date_standard.empty?
				date_standard = date.gsub(/[^0-9]/ , '').slice(0..3)
			end
			if date_standard != date
				date_narrative = date
			end
		end
		# ----------------------------------------------
		contributor = ''
		contributor_rec = record_section.scan(/cet">([^=]*)[.,]<\/a> *[ps][a-z]*r/)
											# Nokogiri
		if contributor_rec.length > 1
			contributor_rec.each do |co|
				contributor += "#{co[0]}. | "
			end
			contributor.gsub! '.. |' , '. |'
			3.times do
				contributor.chop!
			end
		elsif contributor_rec.length > 0
			contributor = contributor_rec[0][0] + '.'
		end
		if contributor[-2] == '.'
			contributor.chop!
		end
		# ----------------------------------------------
		note = ''
		notifications.each do |n|
			note << "[#{n}]"
		end
		# ----------------------------------------------
		[
			call_num, 				  # 0
			url, 							  # 1
			current_repository, # 2
			current_collection, # 3
			current_location,   # 4
			author,						  # 5
			title,							# 6
			place_of_pub,			  # 7
			date_narrative,		  # 8
			date_standard,			# 9
			contributor,				# 10
			note                # 11
		]
	end

	def configure_rows(count)
		config = [10, 5, 6, 7, 9, 13, 14, 15, 16, 17, 18, 4]
		if count == 0
			config
		else
			offset = 42 + (16 * (count - 1))
			new_config = []
			config.each do |row|
				new_config << (row + offset)
			end
			new_config
		end
	end

	def enter_data sheet, column, config, fields
		unless fields.empty?
			(0...fields.length).each do |field_num|
				row = config[field_num]
				sheet[0].add_cell(row - 1, column - 1, fields[field_num].to_s)
			end
		end
	end
end

# =========================================================================================

require 'open-uri'
require 'pathname'
require 'roo'
require 'rubyXL'

c = CatScraper.new

file  		  = c.get_spreadsheet
read_sheet  = Roo::Spreadsheet.open(file).sheet(0)
write_sheet = RubyXL::Parser.parse(file)
tracker     = {}
range 			= c.get_columns

puts ''

range.each do |column|
	call_num = read_sheet.cell(10, column)
	heading  = c.int_to_alph(column)
	print "[#{heading}]: "

	# skip if there's nothing in the call number field
	if call_num.nil?
	  puts 'none'
		next
	else
		call_num.strip!
	end

	# check if there's already data in the column and allow the user to skip it
	skip_check = c.check_cells(read_sheet, column, heading)
	if skip_check
		next
	end

	# use stored information if the call number has already been searched
	if tracker.include? call_num
		puts "\" \"			(#{call_num})"
		rows = c.configure_rows(0)
		c.enter_data(write_sheet, column, rows, tracker[call_num])
	else
		# search the catalog with the call number
		print call_num
		search_url 			= "http://franklin.library.upenn.edu/search.html?q=%22#{call_num.gsub(' ' , '%20')}%22&meta=t"
		results 				= open(search_url).read.scan(/id=FRANKLIN_([0-9]*)/).uniq
		if results.empty?
			puts "		--[NO RESULTS]--"
			tracker[call_num] = []
			next
		end
		print "\n"

	# for each result, copy the records from the catalog page and add them to the spreadsheet
		record_count    = 0
		no_more_records = false
		notifications		= []
		until no_more_records

			if results.length > 1
				notifications[0] = "RESULT ##{record_count + 1}"
			end

		  frankl_id = results[record_count][0]
			cat_url   = "http://franklin.library.upenn.edu/record.html?id=FRANKLIN_#{frankl_id}"
			record    =	open(cat_url).read.scan(/<section id="recordsection.*<div id="gotovshelf">/m)[0].gsub('&amp;' , '&')
			fields    = c.get_fields(call_num, cat_url, record, notifications)

			# add fields to tracker
			unless tracker.include? call_num
				tracker[call_num] = fields
			end

			# records after the first are entered lower in the excel column
			rows = c.configure_rows(record_count)
			c.enter_data(write_sheet, column, rows, fields)

			record_count += 1
			if record_count == results.length
				no_more_records = true
			end

		end
	end
end

write_sheet.write
puts 'done'
