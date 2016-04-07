###
# This program reads the JSON specification files and creates the Markdown files (minus the examples). 
# Location: https://github.com/sumurthy/md_apispec
###
require 'pathname'
require 'logger'
require 'json'

module SpecMaker

	# Initialize 
	JSON_SOURCE_FOLDER = "../jsonFiles/js"	
	#JSON_SOURCE_FOLDER =  "C:/Users/suramam/git/wip/lib/jsonFiles"
	ENUMS = '../jsonFiles/settings/enums.json'
	MARKDOWN_OUTPUT_FOLDER = "../jsOutputMarkdowns/"
	EXAMPLES_FOLDER = "../js-api-examples-to-merge/"
	HEADERQUALIFIER = " Object (JavaScript API for Excel)"
	APPLIESTO = "_Applies to: Excel 2016, Office 2016_"
	HEADER1 = '# '
	HEADER2 = '## '
	HEADER3 = '### '
	HEADER4 = '#### '
	HEADER5 = '##### '
	GETTERSETTERLINK = '_See property access [examples.](#property-access-examples)_'
	GETTERSETTER = 'Property access examples'
	# GETTER = 'Getter Examples'
	# SETTER = 'Setter Examples'
	BACKTOMETHOD = '[Back](#methods)'
	NEWLINE = "\n"
	BACKTOPROPERTY = NEWLINE + '[Back](#properties)'
	PIPE = '|'
	TWONEWLINES = "\n\n"
	PROPERTY_HEADER = "| Property	   | Type	|Description" + NEWLINE
	TABLE_2ND_LINE =  "|:---------------|:--------|:----------|" + NEWLINE
	PARAM_HEADER = "| Parameter	   | Type	|Description|" + NEWLINE
	TABLE_2ND_LINE_PARAM =  "|:---------------|:--------|:----------|" + NEWLINE

	RELATIONSHIP_HEADER = "| Relationship | Type	|Description|" + NEWLINE
	METHOD_HEADER = "| Method		   | Return Type	|Description|" + NEWLINE
	SIMPLETYPES = %w[int string object object[][] double bool number void object[]]

	# Log file
	LOG_FOLDER = '../../logs'
	Dir.mkdir(LOG_FOLDER) unless File.exists?(LOG_FOLDER)

	if File.exists?("#{LOG_FOLDER}/#{$PROGRAM_NAME.chomp('.rb')}.txt")
		File.delete("#{LOG_FOLDER}/#{$PROGRAM_NAME.chomp('.rb')}.txt")
	end
	@logger = Logger.new("#{LOG_FOLDER}/#{$PROGRAM_NAME.chomp('.rb')}.txt")
	@logger.level = Logger::DEBUG
	# End log file

	# Create markdown folder if it doesn't already exist
	Dir.mkdir(MARKDOWN_OUTPUT_FOLDER) unless File.exists?(MARKDOWN_OUTPUT_FOLDER)	

	if !File.exists?(JSON_SOURCE_FOLDER)
		@logger.fatal("JSON Resource File folder does not exist. Aborting")
		abort("*** FATAL ERROR *** Input JSON resource folder: #{JSON_SOURCE_FOLDER} doesn't exist. Correct and re-run." )
	end

	if !File.exists?(EXAMPLES_FOLDER)
		puts "API examples folder does not exist"
	end		

	## 
	# Load up all the known existing enums.
	###
	@enumHash = {}
	begin
		@enumHash = JSON.parse File.read(ENUMS)
	rescue => err
		@logger.warn("JSON Enumeration input file doesn't exist: #{@current_object}")
	end

	@mdlines = []
	@resource = ''
	@gsType = ''

	def self.uncapitalize (str="")
		if str.length > 0
			str[0, 1].downcase + str[1..-1]
		else
			str
		end
	end

	# Write properties and methods to the final array.
	def self.push_property  (prop = {})
		# Add read-only and possible Enum values from the list. 
		
		finalDesc = prop[:isReadOnly] ? prop[:description]  + ' Read-only.' : prop[:description]
		appendEnum = ''
		if (prop[:enumNameJs] != nil) && (@enumHash.has_key? prop[:enumNameJs])
			if @enumHash[prop[:enumNameJs]].values[0] == "" || @enumHash[prop[:enumNameJs]].values[0] == nil
				appendEnum = " Possible values are: " + @enumHash[prop[:enumNameJs]].keys.join(', ') + "."
			else
				appendEnum = " Possible values are: " + @enumHash[prop[:enumNameJs]].map{|k,v| "`#{k}` #{v}"}.join(',') 
			end
			finalDesc = finalDesc + appendEnum
		end
		# If the type is of	an object, then provide markdown link.
		if SIMPLETYPES.include? prop[:dataType] 	
			dataTypePlusLink = prop[:dataType] 	
		else			
			dataTypePlusLink = "[" + prop[:dataType] + "](" + prop[:dataType].downcase + ".md)"
		end

		if prop[:isCollection] 
			dataTypePlusLink = "[" + prop[:dataType] + "](" + prop[:dataType].chomp('[]').downcase + ".md)"
		end
			
		@mdlines.push (PIPE + prop[:name] + PIPE + dataTypePlusLink + PIPE + finalDesc + PIPE) + NEWLINE
	end

	# Write methods to the final array.
	def self.push_method (method = {})

		# If the type is of	an object, then provide markdown link.
		if SIMPLETYPES.include? method[:returnType]
			dataTypePlusLink = method[:returnType]
		else			
			dataTypePlusLink = "[" + method[:returnType] + "](" + method[:returnType].downcase + ".md)"
		end
		# Add anchor links to method. 
		str = method[:signature].strip
		replacements = [ [" ", "-"], ["[", ""], ["]", ""],["(", ""], [")", ""], [",", ""], [":", ""] ]				
		replacements.each {|replacement| str.gsub!(replacement[0], replacement[1])}
		methodPlusLink = "[" + method[:signature].strip + "](#" + str.downcase + ")"
		@mdlines.push (PIPE + methodPlusLink + PIPE + dataTypePlusLink + PIPE + method[:description] + PIPE) + NEWLINE
	end

	# Write methods details and parameters to the final array.	
	def self.push_method_details (method = {}, examples = [])

		@mdlines.push NEWLINE + HEADER3 + method[:signature] + NEWLINE	
		@mdlines.push method[:description] + TWONEWLINES	
		@mdlines.push HEADER4 + "Syntax" + NEWLINE + '```js' + NEWLINE
		@mdlines.push method[:syntax] + NEWLINE + '```' + TWONEWLINES
		@mdlines.push HEADER4 + "Parameters" + NEWLINE

		if method[:parameters] !=nil  			

			@mdlines.push PARAM_HEADER + TABLE_2ND_LINE_PARAM 
			method[:parameters].each do |param|
				# Append optional and enum possible values (if applicable).
				finalPDesc = param[:isRequired] ? param[:description] : 'Optional. ' + param[:description]
				appendEnum = ''
				if (param[:enumNameJs] != nil) && (@enumHash.has_key? param[:enumNameJs])

					if @enumHash[param[:enumNameJs]].values[0] == "" || @enumHash[param[:enumNameJs]].values[0] == nil
						appendEnum = " " + " Possible values are: " + @enumHash[param[:enumNameJs]].keys.join(', ')  
					else
						appendEnum = " Possible values are: " + @enumHash[param[:enumNameJs]].map{|k,v| "`#{k}` #{v}"}.join(',')
					end
					finalPDesc = finalPDesc + appendEnum
				end
				@mdlines.push (PIPE + param[:name] + PIPE + param[:dataType] + PIPE + finalPDesc + PIPE) + NEWLINE	
			end
		else
			@mdlines.push "None"  + NEWLINE
		end

		@mdlines.push NEWLINE + HEADER4 + "Returns" + NEWLINE

		if SIMPLETYPES.include? method[:returnType]
			dataTypePlusLink = method[:returnType]
		else			
			dataTypePlusLink = "[" + method[:returnType] + "](" + method[:returnType].downcase + ".md)"
		end
		@mdlines.push dataTypePlusLink + NEWLINE
		

		# loc:100
		if	@exampleFileFound == true
			exampleFound	 = false
			examples.each_with_index do |exampleLine, i|
				if (exampleLine.chomp.strip.include? method[:name]) && (exampleLine.chomp.strip.include?('###'))
					exampleFound = true
				# moving here from loc:100	
					@mdlines.push NEWLINE + HEADER4 + 'Examples' + NEWLINE
				# end move
					next
				end

				if exampleFound && exampleLine.start_with?('##')
					break
				end
				if exampleFound	 
					@mdlines.push exampleLine
				end
			end
			# comment below 5 lines to not print empty example block when the example is not found. 
			# if !exampleFound
			# 	@mdlines.push "```js" + TWONEWLINES
			# 	@mdlines.push "```" + NEWLINE
			# 	@logger.error("....Example not found for method: #{method[:signature]}, #{@resource}  ") 
			# end
		end
		#@mdlines.push NEWLINE + BACKTOMETHOD + TWONEWLINES 
		
	end

	# Add getter and setter examples
	def self.push_getter_setters (examples = [] )
		getterOrSetterFound	 = false

		examples.each_with_index do |exampleLine, i|
			if (exampleLine.chomp.strip.downcase.include? "getter") || (exampleLine.chomp.strip.downcase.include? "setter")
				getterOrSetterFound = true
					@mdlines.push HEADER3 + GETTERSETTER + NEWLINE 
				next
			end
			if getterOrSetterFound && exampleLine.include?('##')
				break
			end
			if getterOrSetterFound	 
				@mdlines.push exampleLine
			end
		end
		# if getterOrSetterFound 
		# 	@mdlines.push BACKTOPROPERTY + NEWLINE
		# end
	end

	# Determine the type getter and setter links to be used. 
	def self.determine_getter_setter_type (examples = [])
		gsType = 'none'
		examples.each_with_index do |exampleLine, i|
			if (exampleLine.chomp.strip.downcase.include? "getter") || (exampleLine.chomp.strip.downcase.include? "setter")
				if (exampleLine.chomp.strip.downcase.include? "getter") && (exampleLine.chomp.strip.downcase.include? "setter")
					gsType = 'getterandsetter'
				elsif (exampleLine.chomp.strip.downcase.include? "getter") 
					gsType = 'getter'	
				else
					gsType = 'setter'
				end
			end
		end
		gsType
	end

	# Conversion to specification 
	def self.convert_to_spec (item=nil)
		@mdlines = []
		@jsonHash = JSON.parse(item, {:symbolize_names => true})
		# Obtain the resource name. Read the examples file, if it exists. 
		@resource = uncapitalize(@jsonHash[:name])
		@log#ger.debug("")	
		@logger.debug("...............Report for: #{@resource}...........")	

		example_lines = ''
		@gsType = ''
		@exampleFileFound = false
		begin
			#example_lines = File.readlines(File.join(JSON_EXAMPLE_FOLDER + @resource.downcase + ".md"))
			example_lines = File.readlines(EXAMPLES_FOLDER + @resource.downcase + ".md")
			@gsType = determine_getter_setter_type example_lines
			@exampleFileFound = true
		rescue => err
			puts "....Example File does not exist for: #{@resource}"
		end

		propreties = @jsonHash[:properties]
		if propreties 
			propreties = propreties.sort_by { |v| v[:name] }
		end

		methods = @jsonHash[:methods]
		if methods 
			methods = methods.sort_by { |v| v[:name] }
		end

		header_name = @jsonHash[:isCollection] ? "List #{@jsonHash[:collectionOf]}" : "Get #{@jsonHash[:name]}"
		@mdlines.push HEADER1 + @jsonHash[:name] + HEADERQUALIFIER + TWONEWLINES
		@mdlines.push  APPLIESTO + TWONEWLINES
		@mdlines.push @jsonHash[:description] + TWONEWLINES

		isRelation, isProperty, isMethod = false, false, false 

		if propreties != nil
			propreties.each do |prop|
				
				if !prop[:isRelationship]
				   isProperty = true
				end

#				puts " #{@resource}..... #{prop[:name]} ..  #{prop["isrelationship"]}... #{prop[:isCollection]} .. #{prop[:description]}"
				if prop[:isRelationship]			  
				   isRelation = true
				end
			end
		end

		if methods != nil
			isMethod = true
		end

		@logger.debug("....Is there: property?: #{isProperty}, relationship?: #{isRelation}, method?: #{isMethod} ..........")	

		# Add property table. 	

		# Add properties header
		@mdlines.push HEADER2 + 'Properties' + TWONEWLINES
		if isProperty
			# add properties table
			@mdlines.push PROPERTY_HEADER + TABLE_2ND_LINE 
			propreties.each do |prop|
				if !prop[:isRelationship]
					@logger.debug("....Processing property: #{prop[:name]} ..........")	
				   push_property prop
				end
			end
			# Sep-20, Property read-write example addition
			if @gsType != 'none'
				@mdlines.push NEWLINE + GETTERSETTERLINK + NEWLINE
			end

		else
			@mdlines.push "None"  + NEWLINE
		end		

		# Add Relationship table. 
		@mdlines.push NEWLINE
		@mdlines.push HEADER2 + 'Relationships' + NEWLINE


		if isRelation
			@mdlines.push RELATIONSHIP_HEADER + TABLE_2ND_LINE 
			propreties.each do |prop|
				if prop[:isRelationship]
					@logger.debug("....Processing relationship: #{prop[:name]} ..........")		
				   push_property prop
				end
			end
		else
			@mdlines.push "None"  + TWONEWLINES
		end		

		# Add method table. 
		@mdlines.push NEWLINE + HEADER2 + 'Methods' + NEWLINE

		if isMethod
			@mdlines.push NEWLINE + METHOD_HEADER + TABLE_2ND_LINE 
			methods.each do |mtd|
				@logger.debug("....Processing method: #{mtd[:name]} ..........")						
				push_method mtd
			end
		else
			@mdlines.push "None"  + TWONEWLINES
		end	

		# Add each API method details.	
		if isMethod || (@gsType != 'none' && @gsType != '') 
			@mdlines.push NEWLINE + HEADER2 + 'Method Details' + TWONEWLINES
		end	

		if isMethod
			methods.each do |mtd|
				push_method_details mtd, example_lines
			end
			
		end
		if @gsType != 'none' && @gsType != '' 
			push_getter_setters example_lines
		end

		# Write the output file. 
		outfile = MARKDOWN_OUTPUT_FOLDER + @resource.downcase + '.md'
		file=File.new(outfile,'w')
		@mdlines.each do |line|
			file.write line
		end
	end

	# Main loop. 
	processed_files = 0
	Dir.foreach(JSON_SOURCE_FOLDER) do |item|
		next if item == '.' or item == '..'
		fullpath = JSON_SOURCE_FOLDER + '/' + item.downcase

		if File.file?(fullpath)
			convert_to_spec File.read(fullpath)
			processed_files = processed_files + 1
		end
	end
	puts ""
	puts "*** OK. Processed #{processed_files} input files. Check #{File.expand_path(LOG_FOLDER)} folder for results. ***"
end