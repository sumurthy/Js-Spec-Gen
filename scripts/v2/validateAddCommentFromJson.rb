###
# This program reads the .CS metadata file and inserts the comments from the JSON specification files.
# Location: https://github.com/sumurthy/md_apispec
###

require 'logger'
require 'json'

module SpecMaker

	LOG_FILE = '../../logs/validateAddComments_log.txt'
	begin
		File.delete(LOG_FILE)
		File.delete(EXCELAPI_FILE_TRANSIT)
	rescue => err
	    #Ignore this error
	end

	@logger = Logger.new(LOG_FILE)
	@logger.level = Logger::DEBUG

	EXCELAPI_FILE_SOURCE = '../../data/ExcelAPI.cs'
	EXCELAPI_FILE_TRANSIT = '../../data/ExcelAPI_transit.cs'
	EXCELAPI_WITH_COMMENTS = '../../data//ExcelAPI_out.cs'
	
	RESOURCE_FOLDER = 'jsonFiles/'

	@csarray = []
	@csarray_pure = []
	@csarray_out = []
	@resource_md = []
	@current_object = ''
	@resource = {}
	@jsonHash = {}

	def self.csarray_write (line=nil) 
		@csarray_out.push line
	end

	def self.parse_resource (item=[]) 
		@resource = {}
		@resource[:dsc] = ''
		@resource[:prp] = ''
		@resource[:mtd] = ''
		@resource[:prm] = ''
		properties_array = []
		methods_array = []
		parameters_array =[]
		
		@jsonHash = JSON.parse(item, {:symbolize_names => true})	

		@resource[:dsc] = @jsonHash[:description]

		propreties_hash = @jsonHash[:properties]

		if propreties_hash != nil
			propreties_hash.each_with_index do |property,  i|
				finalDescription = property[:description]
				if property[:enumNameJs] != nil
					finalDescription = finalDescription + " See " + property[:enumNameJs] + " for details."
				end 
				properties_array.push [property[:name], property[:dataType], finalDescription]
			end
		end

		methods_hash = @jsonHash[:methods]

		if methods_hash != nil
			methods_hash.each_with_index do |method,  i|
				methods_array.push [method[:name], method[:returnType], method[:description]]
			end
		end

		if methods_hash != nil  			
			methods_hash.each do |method|
				if method[:parameters] != nil
					method[:parameters].each do |param|
						finalDescription = param[:description]
						if param[:enumNameJs] != nil
							finalDescription = finalDescription + " See " + param[:enumNameJs] + " for details."
						end
						parameters_array.push [param[:name], param[:dataType], finalDescription, method[:name] ]	
					end
				end
			end
		end

		@resource[:prp] = properties_array 
		@resource[:mtd] = methods_array
		@resource[:prm] = parameters_array
	# End of method
	end

	# Fill description 

	def self.fill_description 
		padding = "\t"
		csarray_write padding + "/// <summary>\n"
		csarray_write padding + "/// " + @resource[:dsc].strip + "\n"
		csarray_write padding + "/// </summary>" + "\n"
	end

	def self.fill_prop (desc="")
		padding = "\t\t"
		csarray_write padding + "/// <summary>\n"
		csarray_write padding + "/// " + desc.strip + "\n"
		csarray_write padding + "/// </summary>" + "\n"
	end

	def self.fill_method (desc="")
		padding = "\t\t"
		csarray_write padding + "/// <summary>\n"
		csarray_write padding + "/// " + desc.strip + "\n"
		csarray_write padding + "/// </summary>" + "\n"

	end

	def self.fill_params (parm_hash = {})
		padding = "\t\t"
		parm_hash.each do |key, val|
			csarray_write padding + '/// <param name="' + key + '">' + val + '</param>' + "\n"
		end
	end

	### 
	# Read the file & create a transit file by removing existing comments from the .CS File.
	##
	csarray_transit = File.readlines(EXCELAPI_FILE_SOURCE) 
	filter_start = false
	handle_getItem = ''
	csarray_transit.each do |line|

		if line.start_with?('#region Application')
			filter_start = true
		end
		#region Enums
		if line.start_with?('#region Enums')
			filter_start = false
		end

		if filter_start && line.include?('///')
			next
		else
			@csarray_pure.push line
		end
	end

	@csarray_pure.each do |line|

		if line.include?('this[')
				handle_getItem = line
				handle_getItem = handle_getItem[0,handle_getItem.index('{')]
				handle_getItem = handle_getItem.gsub('this[','getItem(').gsub(']',');')
				line = handle_getItem + "\n"
		end
		@csarray.push line
		
	end

	### 
	# Just for referencing purpose: create a transit file by removing existing comments from the .CS File. 
	##
	file=File.new(EXCELAPI_FILE_TRANSIT,'w')

	@csarray.each do |line|
	    file.write line
	end


	#@csarray = File.readlines(EXCELAPI_FILE) 



	### 
	# Forward Pass: Write to the output array
	##

	@csarray.each_with_index do |line, i|

		## For new object, load its resource and fill the description
		if line.strip.start_with?('[ClientCallableComType', '[ClientCallableServiceRoot', '[ClientCallableType') && \
		   !@csarray[i-1].strip.start_with?('[ClientCallableComType', '[ClientCallableServiceRoot', '[ClientCallableType')

			j = i
			until (@csarray[j].include?('public interface'))
					j = j+1
			end

			if j > i
				temp = @csarray[j].split.first(3).join(' ')

				# Get the third word
				@current_object = temp.split.last(1).join(' ').gsub(':','')

				begin
					@resource_json = File.read(RESOURCE_FOLDER + @current_object.downcase + ".json")
				rescue => err
				  @logger.fatal("Resource File does not exist for: #{@current_object}")
				  @logger.fatal(err)
				  abort("*** FATAL ERROR *** #{@current_object}")
				end

				# Send the JSON file for processing so that it could be loaded up for processing.
				parse_resource(@resource_json)
				# Print Logs
					@logger.debug("...............Report for: #{@current_object}...........")
					@logger.debug("Description: #{@resource[:dsc]}")
					@logger.debug("Properties: #{@resource[:prp].length}")
					@resource[:prp].each do |line|  
						@logger.debug("#{line}")
					end	
					@logger.debug("Methods: #{@resource[:mtd].length}")
					@resource[:mtd].each do |line| 
						@logger.debug("#{line}")
					end	
					@logger.debug("Params: #{@resource[:prm].length}")
					@resource[:prm].each do |line|
						@logger.debug("#{line}")
					end	
				# END Print Logs
			end		
			fill_description
		end

		## For each property or method, fill its comments
		if line.strip.start_with?('[ClientCallableComMember', '[ClientCallableOperation') && \
		   !@csarray[i-1].strip.start_with?('[ClientCallableComMember', '[ClientCallableOperation')
			j = i			
			while (@csarray[j].include?('[ClientCallable'))
					j = j+1
			end
			# Process only if it is not an internal property/method
			if !@csarray[j].include?('_')
				# Presence of { would indicate that it is a property or a relation	
				#if @csarray[j].include?('{') && !@csarray[j].include?('this[')
				if @csarray[j].include?('{')  
					prop_name = @csarray[j].split[1]
					match_found = false
					k=0
					while !match_found && k < @resource[:prp].length do 
						if @resource[:prp][k][0].casecmp(prop_name) == 0
							match_found = true
							# Insert Read-Only if setter is missing.					
							if setter = @csarray[j].include?('set;') 
								descToSend = @resource[:prp][k][2] 
							else
								descToSend = @resource[:prp][k][2] + " Read-only."
							end
							fill_prop (descToSend)
						end
						k = k + 1
					end
					if !match_found && !prop_name.start_with?('_')
						@logger.warn("......No Properties Match for: #{@current_object} . #{prop_name}..........")
					end

				# Special case the getItem
				# elsif @csarray[j].include?('this[')
					#Do nothing for now		
				# Remaining is the method call	
				else
					temp = @csarray[j].split[1]
					mthd_name = temp[0,temp.index('(')]
					match_found = false
					k=0

					while !match_found && k < @resource[:mtd].length do 
						if @resource[:mtd][k][0].casecmp(mthd_name) == 0
							match_found = true
							fill_method @resource[:mtd][k][2]
						end
						k = k + 1
					end
					if !match_found && !mthd_name.start_with?('_')
						@logger.warn("..No Method Match for: #{@current_object} . #{mthd_name}..........")
					end
					# if method has params, then load them in an array and add <param.. comments
					if !@csarray[j].include?('();')
						#parm_array = @csarray[j][@csarray[j].index('('), @csarray[j].index(')')].split(',')
						parm_array = @csarray[j][@csarray[j].index('('), @csarray[j].index(');')].split(',')
						parm_array.map! {|n| n.split[1]}
						parm_array.map! {|n| n.gsub(');','')}					
						
						param_hash = {}

						parm_array.each do |val|
							param_hash[val] = ''
						end

						parm_array.each_with_index do |val, pos|
							@resource[:prm].each do |prmarray|
								if prmarray[0].casecmp(val) == 0 
									if prmarray[3].casecmp(mthd_name) == 0 
										param_hash[val] = prmarray[2]
										break	
									end
								end
							end
							if param_hash[val] == ''
								@logger.warn("..No Param Match for: #{@current_object} . #{val}..........")			
							end
						end
						fill_params param_hash

					end
				end
			end
		end
		csarray_write @csarray_pure[i]
	end

	##
	# Write the CS Out file
	##

	file=File.new(EXCELAPI_WITH_COMMENTS,'w')

	@csarray_out.each do |line|
	    file.write line
	end
end