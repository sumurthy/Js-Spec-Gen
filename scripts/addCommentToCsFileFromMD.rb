###
# This program reads the .CS metadata file and inserts the comments from the Markdown specification files.
# Location: https://github.com/sumurthy/wip
###


require 'logger'

begin
	File.delete('logs/logfile.txt')
	File.delete(EXCELAPI_FILE_TRANSIT)
rescue => err
    #Ignore this error
end

@logger = Logger.new('logs/logfile.txt')
@logger.level = Logger::DEBUG

@logger.debug("Created logger")
@logger.info("Program started")
@logger.warn("Nothing to do!")

EXCELAPI_FILE_SOURCE = '../ExcelAPI.cs'
EXCELAPI_FILE_TRANSIT = '../ExcelAPI_transit.cs'
EXCELAPI_WITH_COMMENTS = '../ExcelAPI_out.cs'
RESOURCE_FOLDER = 'C:\Users\suramam\git\ExcelJsDocumentation\resources'

@csarray = []
@csarray_pure = []
@csarray_out = []
@resource_md = []
@current_object = ''
@resource = {}

def csarray_write (line=nil) 
	@csarray_out.push line
end

def parse_resource (lines=[]) 
	desc = false
	@resource = {}
	properties = []
	methods = []
	params = []
	@resource[:dsc] = ''
	@resource[:prp] = ''
	@resource[:mtd] = ''
	@resource[:prm] = ''

	# start Super Loop
	#lines.each_with_index do |line, i|
	i, j, k, l = 0, 0, 0, 0
	while i < lines.length 

		if i == 0 || lines[i] == ''
			i = i + 1
			next
		end
		#if lines[i].start_with?('## API Specification')
		#	break
		#end
		if lines[i].start_with?('## Properties', '## [Properties',\
							 '## Relationship', '## [Relationship',\
							 '## Methods')
			desc = true
			
		end
		# Concatenate all descriptions and remove new-line from the object description.
		if i >= 1 && !desc
			@resource[:dsc] << " #{lines[i].strip.gsub(/\s+/, ' ')}"
		end

		# Properties Search and load name/type/description
		if lines[i].start_with?('## Properties', '## [Properties')
			until (lines[i].start_with?(':--', '|:--', 'None'))
				i = i+1
			end
			if (lines[i].start_with?(':--', '|:--'))
				i = i+1
				while lines[i].start_with?('|') 
					properties[j] = lines[i].split('|')[1,3]
					properties[j].map!(&:strip)
					properties[j].map! {|n| n.gsub('`','')}
					i = i + 1
					j = j + 1
				end				
			end
		# End Properties Search
		end

		# Relationship Search and load name/type/description
		if lines[i].start_with?('## Relationship', '## [Relationship')
			until (lines[i].start_with?(':--', '|:--', 'None'))
				i = i+1
			end
			if (lines[i].start_with?(':--', '|:--'))
				i = i+1
				while lines[i].start_with?('|') 
					properties[j] = lines[i].split('|')[1,3]
					properties[j].map!(&:strip)
					properties[j].map! {|n| n.gsub('`','')}
					i = i + 1
					j = j + 1
				end				
			end
		# End Relationship Search
		end

		# Methods Search and load name/type/description
		if lines[i].start_with?('## Method', '## [Method')
			until (lines[i].start_with?(':--', '|:--', 'None'))
				i = i+1
			end
			if (lines[i].start_with?(':--', '|:--'))
				i = i+1
				while lines[i].start_with?('|') 
					methods[k] = lines[i].split('|')[1,3]
					methods[k].map!(&:strip)
					methods[k].map! {|n| n.gsub('`','')}
					i = i + 1
					k = k + 1
				end				
			end
		# End Methods Search
		end

		#Save the method name to be loaded into the Param table for later comparison with the .cs file. Some params repeat in the same resource and they have different meanings..
		#Extract only the method name
		if lines[i].start_with?('### ')
			if lines[i].include?('(') 
				method_name_save = lines[i][4,lines[i].index('(')-4]
			else
				method_name_save = 'INVALID'
			end
		end
		# Parameter Search and load name/type/description
		if lines[i].start_with?('#### Parameter') || lines[i].strip.start_with?('Parameter|')
			until (lines[i].start_with?(':--', '|:--', '--', 'None'))
				i = i+1
			end
			if (lines[i].start_with?(':--', '|:--', '--'))
				i = i+1
				while lines[i].include?('|')
					if lines[i].start_with?('|')
						params[l] = lines[i].split('|')[1,3]
					else
						params[l] = lines[i].split('|')[0,3]
					end
					params[l].map!(&:strip)
					params[l].map! {|n| n.gsub('`','')}
					params[l][3] = method_name_save
					i = i + 1
					l = l + 1
				end				
			end
		# End Methods Search
		end
		# Increment Resource MD file counter. 
		i = i + 1
	# end Super Loop
	end

	# Clean up the items to remove markdown links.
	properties.each_with_index do |array, index|

		array.each_with_index do |item, innerindex|
			if innerindex < 2 && item.index('(') != nil
				item = item[0,item.index('(')]
				properties[index][innerindex] = item	
			end
		end 	
		properties[index].map! {|n| n.gsub('[','')}
		properties[index].map! {|n| n.gsub(']','')}		
	end 

	methods.each_with_index do |array, index|

		array.each_with_index do |item, innerindex|
			if innerindex < 2 && item.index('(') != nil 
				item = item[0,item.index('(')]
				methods[index][innerindex] = item	
			end
		end 	
		methods[index].map! {|n| n.gsub('[','')}
		methods[index].map! {|n| n.gsub(']','')}	
	end 


	params.each_with_index do |array, index|

		array.each_with_index do |item, innerindex|
			if innerindex < 2 && item.index('(') != nil
				item = item[0,item.index('(')]
				params[index][innerindex] = item	
			end
		end 	
		params[index].map! {|n| n.gsub('[','')}
		params[index].map! {|n| n.gsub(']','')}

	end 

	# Load the resource hash with its members. 
	# @resource[:dsc] already has the description loaded. 
	@resource[:prp] = properties
	@resource[:mtd] = methods
	@resource[:prm] = params

# end of method
end

# Fill description 

def fill_description 
	padding = "\t"
	csarray_write padding + "/// <summary>\n"
	csarray_write padding + "/// " + @resource[:dsc].strip + "\n"
	csarray_write padding + "/// </summary>" + "\n"
end

def fill_prop (desc="")
	padding = "\t\t"
	csarray_write padding + "/// <summary>\n"
	csarray_write padding + "/// " + desc.strip + "\n"
	csarray_write padding + "/// </summary>" + "\n"
end

def fill_method (desc="")
	padding = "\t\t"
	csarray_write padding + "/// <summary>\n"
	csarray_write padding + "/// " + desc.strip + "\n"
	csarray_write padding + "/// </summary>" + "\n"

end

def fill_params (parm_hash = {})
	padding = "\t\t"
	parm_hash.each do |key, val|
		csarray_write padding + '/// <param name="' + key + '">' + val + '</param>' + "\n"
	end
end
#                                /// <summary>
#                                /// 
#                                /// </summary>
#                                /// <param name="row"></param>
#                                /// <param name="column"></param>

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
				@resource_md = File.readlines(File.join(RESOURCE_FOLDER, @current_object+".md"))

			rescue => err
			  @logger.fatal("Resource File does not exist for: #{@current_object}")
			  @logger.fatal(err)
			  abort("*** FATAL ERROR *** #{@current_object}")
			end

			parse_resource(@resource_md)
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