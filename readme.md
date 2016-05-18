# Office JavsScript Spec Maker

This tool is designed to generate Office JavaScript API specification docs based on internal definition/design. 

Running this tool involves 3 things: 

1. Setup the input data as described below. 
2. Run intermediary JSON file creation step (`genJsonFromMetadata.rb` script). 
3. Run the markdown creation step (`genMarkdownFromJSON.rb` script), which produces the final API spec files. You can copy these files into final destination such as another repository. 

## Pre-requisites

Ruby interpreter. Version 2.1+ 

## Input data 

### Metadata definition file

This is an internal metadata defintion file that describes the API structure. This file contains all of the objects, proeprties, relationships, methods, parameters, etc. along with their description. Additional information such as version supported in, etc. are also provided. 

##### File Location

This file is not part of the repository since it contains some internal data. 
The location of this file is defined in `config/config.json` script file in `metadataFilePath` property. The path is relative path from the script files. If you wish to change the location, just update the location of the source folder. 

**Please do not make source files part of the repository as it doesn't belong here**


## Output markdown spec

The final output files are generated in markdown format. 

##### File Location

The location of final spec files are at `Js-Spec-Gen/markdown`. It includes one file per object. 

## Run steps

1. Setup the input data 
2. change directory to `Js-Spec-Gen/scripts` folder
2. Run Json intermediary file creation step: `ruby genJsonFromMetadata.rb`
3. Once the above step completes, run the markdown creation step. `ruby genMarkdownFromJSON.rb`
4. Find your markdown files in the output folder. 
