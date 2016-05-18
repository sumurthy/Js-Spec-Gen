###
# This program reads the JSON specification files and creates the Markdown files (minus the examples). 
# Location: https://github.com/sumurthy/Js-Spec-Gen
###
require 'pathname'
require 'json'
require 'FileUtils'
require 'base64'

module SpecMaker

	# Initialize 
	NEWLINE = "\n"
	JSON_SOURCE_FOLDER = "jsonFiles/source"		
	ENUMS = 'jsonFiles/settings/enums.json'
	MARKDOWN_OUTPUT_FOLDER = "../markdown/"
	WRITE_BACK = %w[# | *]
	TAKE_ACTION = %w[% >]
	IGNORE = %w[/]

	EXAMPLES_FOLDER = "../api-examples-to-merge/"

	SIMPLETYPES = %w[int string object object[][] double bool flaot number void object[]]

	# Read config and json_struct files 

	CONFIG = "../config/config.json"
	@config = JSON.parse(File.read(CONFIG, :encoding => 'UTF-8'), {:symbolize_names => true})
	puts "....Starting run for the app #{@config[:app]}"
	puts
	## 
	# Load the output template 
	###
	@md_main = []
	@md_method = []
	@mdo = []
	@jsonHash = {}
	@region = 'object'

	begin
		@md_main = File.readlines(@config[:mdTemplateMain])
	rescue => err
		abort("*** FATAL ERROR *** Input MD template file: #{@config[:mdTemplateMain]} doesn't exist. Correct and re-run." )
	end

	begin
		@md_method = File.readlines(@config[:mdTemplateMethod])
	rescue => err
		abort("*** FATAL ERROR *** Input MD template file: #{@config[:mdTemplateMethod]} doesn't exist. Correct and re-run." )
	end

	# Create markdown folder if it doesn't already exist
	Dir.mkdir(MARKDOWN_OUTPUT_FOLDER) unless File.exists?(MARKDOWN_OUTPUT_FOLDER)	

	# Clean-up the markdown folder
	FileUtils.rm Dir.glob(MARKDOWN_OUTPUT_FOLDER + '/*')

	if !File.exists?(JSON_SOURCE_FOLDER)
		abort("*** FATAL ERROR *** Input JSON resource folder: #{JSON_SOURCE_FOLDER} doesn't exist. Correct and re-run." )
	end

	if !File.exists?(EXAMPLES_FOLDER)
		puts "API examples folder does not exist"
	end		


	def self.decode(desc="")
		return Base64.decode64(desc).split('|').join(" 	 \n")
	end	

	def self.hyperlink
	end

	def self.substitute(line="")		
		(line.sub! '%resourcename%', @jsonHash[:name]) if line.include?('%resourcename%')
		(line.sub! '%resourcedescription%', @jsonHash[:description]) if line.include?('%resourcedescription%')
		(line.sub! '%longobjectdescription%', (decode @jsonHash[:longDesc])) if line.include?('%longobjectdescription%')
		return line
	end

	def self.process_param
	end

	def self.process_properties
	end

	def self.process_enums
	end

	def self.process_method_details
	end

	def self.process_object
	end

	def self.direct(key='', key2= '', val='')
		if key == '%'
			val = substitute val
		end
		return val
	end

	# Conversion to specification 
	def self.convert_to_spec (item=nil)
		@jsonHash = JSON.parse(item, {:symbolize_names => true})
		@region = 'object'

		# Obtain the resource name. Read the examples file, if it exists. 
		@resource = @jsonHash[:name]
		
		# example_lines = ''
		# @exampleFileFound = false
		# begin
		# 	#example_lines = File.readlines(File.join(JSON_EXAMPLE_FOLDER + @resource.downcase + ".md"))
		# 	example_lines = File.readlines(EXAMPLES_FOLDER + @resource.downcase + ".md")
		# 	@gsType = determine_getter_setter_type example_lines
		# 	@exampleFileFound = true
		# rescue => err
		# 	puts "....Example File does not exist for: #{@resource}"
		# end

		@md_main.each_with_index do |tline, i|
			key = tline.to_s[0]
			key2 = tline.to_s[0..1]
			key = '*' if key.strip.length == 0
			val = tline.strip

			hasVar = val.include?('%') ? true  : false

			case key 
			when *WRITE_BACK
				val = substitute(val) if hasVar
				@mdo.push val + NEWLINE
				next
			when *TAKE_ACTION
				@mdo.push (direct(key, key2, val) + NEWLINE)
				next
			when *IGNORE
				next
			else
				next
			end

		end
	end

	# Main loop. 
	processed_files = 0

	Dir.foreach(JSON_SOURCE_FOLDER) do |item|
		next if item == '.' or item == '..'
		fullpath = JSON_SOURCE_FOLDER + '/' + item.downcase

		if File.file?(fullpath)
			convert_to_spec File.read(fullpath)
		end

		outfile = MARKDOWN_OUTPUT_FOLDER + item.chomp('.json') + '.md'

		file=File.new(outfile,'w')
		@mdo.each do |line|
			file.write line
		end
		processed_files = processed_files + 1

	end
	puts ""
	puts "*** OK. Processed #{processed_files} input files. ***"
end
