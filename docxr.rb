#!/usr/bin/env ruby

# Copyright (c) 2010, James A. Feister, OpenJAF@gmail.com
# All rights reserved.

# Redistribution and use in source and binary forms, with or without 
# modification, are permitted provided that the following conditions are met:
#    * Redistributions of source code must retain the above copyright notice, 
#		this list of conditions and the following disclaimer.
#    * Redistributions in binary form must reproduce the above copyright notice,
#		this list of conditions and the following disclaimer in the 
#		documentation and/or other materials provided with the distribution.
#    * Neither the name of the <ORGANIZATION> nor the names of its contributors
#		 may be used to endorse or promote products derived from this software 
#		without specific prior written permission.

##THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" 
#	AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE 
#	IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE 
#	ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE 
#	LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR 
#	CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF 
#	SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS 
#	INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
#	CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) 
#	ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE 
#	POSSIBILITY OF SUCH DAMAGE.

# Name: docxr.rb
# Description: Will accept a Microsoft .docx Word file and print its contents 
#				to stdout
# Example: 'cat <filename.docx> | ./antiword.rb'
#

############
# requires #
############
# Required for all to work
require 'rubygems'
# Required for proccessing zip files.
require 'zip/zip'
require 'zip/zipfilesystem'
# Required for proccessing the xml.
require 'rexml/document'
# Required for proccessing the stdin
require 'fcntl'
# Required for temp files
require 'tempfile'


# default line wrap width
$termWidth = 80 

# The main document string that contains all raw data in the document.
$docString = ""

#######
# Name: createTempZip
# Desc: Create a temproary version of zip file from stdin
# Preq: stdin pipe of a zip file 
# Post: return temporary zip file descriptor 
def createTempZip( zipFileContents )
        tempdocx = Tempfile.new('antiword')
		fDescriptor = File.new( tempdocx.path, 'w')
        tempdocx.print( zipFileContents )
        tempdocx.close
		return tempdocx
end

#######
# Name: unzipFile
# Desc: Extract requested file from the zip archive
# Preq: zip file name to extract from, file name to get from archive
# Post: returns contents of the file read, or nil on fail
def unzipFile(zipFile, fileName)
	fileContents = ""
	Zip::ZipFile.open(zipFile) { |zip_file|
		fileContents = zip_file.read(fileName) 
		zip_file.close()
	}
	return fileContents
end

#######
# Name: getDocument
# Desc: using the global stdin will return the document.xml from the file.
# Preq: stdInput contains data has been populated
# Post: Returns content of the "word/document" file 
def getDocumentXML #( stdInput )
	# Create a temporary file with a copy of the stdin
	docxFile = createTempZip( $stdin.read )

	# Get the contents of the docx's "word/document.xml"
	document_contents =  unzipFile(docxFile.path, "word/document.xml")

	# Clean up the temporary file
	docxFile.delete

	return document_contents
end

#######
# Name: usage
# Desc: Print out usage information for antiword.rb
# Preq: None
# Post: Usage information is sent to stdout. 
def usage ()
puts %q{
USAGE: 
	cat [ FILE ] | antiword.rb [--help] [-help] [-h] [-t <column width>]

DESCRIPTION: 
	Accepts a .docx formated file printing the text contents to stdout.

REQUIRED FIELDS:
	Piped in file.

OPTIONS:
	--help, -help, -h   : Display a help screen and sample usage.
	-w <integer>        : Set width of text in columns, default is 80

}
end

# Text item
class TextItem
	attr_reader :content
	@item
	def initialize( titem )
		# Internal copy
		@item = titem
		@content = Array.new	
		# Preseve space from the begining and end of line?
		eolSpace = false
		if( @item.attributes["xml:space"] == "preserve" )
			# Does this line only contain space??
			if( @item.text.match(/\S/) == nil )
				#puts "No chars"
				@content.push(@item.text.chop)
				return
			else
				if( @item.text.match(/(^ +)/) )
					@content.push("#{@item.text.match(/(^ +)/)}".chop)
				end
				eolSpace = true # We cant append to the end yet	
			end
		end

		# Add text to the array
		@item.text.strip.scan(/\S+/).each { |e| 
			#puts "Pushing: \"#{e}\""
			@content.push(e) 
		}

		# Append the rest of the End Of line space?
		if( eolSpace == true )
			if (@item.text.match(/( +$)/) )
				@content.push("#{@item.text.match(/( +$)/)}".chop )
			end
		end 
	end
end

# Row item - contained in "w:p"
class RItem
	attr_reader :content
	@item
	def initialize( ritem )
		@item = ritem 
		@content = Array.new()
		# get the text item
		@item.elements.each { |e|
			if( e.expanded_name == "w:t" ) 
				add_content( TextItem.new( e ).content )
			end
		}
	end
	def add_content( contents )
		contents.each { |r|
			@content.push(r)
		}
	end
end

# Hyperlink item
class HLItem
	attr_reader :content
	def initialize( hlitem )
		@item = hlitem 
		@content = Array.new()
		@item.elements.each { |e| 
			if( e.expanded_name == "w:r" )
				add_content(RItem.new( e ).content)
			end
		}
	end
	def add_content( contents )
		contents.each { |c| 
			@content.push( c )
		}
	end
end

# found in paragraphs
class SmartTag
	attr_reader :content
	@item
	@width 
	def initialize( smartTag )
		@item = smartTag
		@content = Array.new()
		@item.elements.each{ |e|
			if( e.expanded_name == "w:r" )
				add_content( RItem.new( e ).content )
			elsif( e.expanded_name == "w:smartTag" )
				add_content( SmartTag.new( e ).content )
			end
		}	
	end
	def add_content( contents )
		contents.each { |c|
			@content.push( c )
		}
	end
end
class ParagraphItem
	attr_reader :content
	@item
	@width
	@tmpline
	def initialize( paragraph , width)
		@item = paragraph
		@width = width
		@content = Array.new()
		@content.push("")
		# format and return text 
		@tmpLine = ""
		@item.elements.each { |e|
			if( e.expanded_name == "w:r" )
				add_content( RItem.new( e ).content )
			elsif( e.expanded_name == "w:hyperlink" )
				add_content( HLItem.new( e ).content )
			elsif( e.expanded_name == "w:smartTag" )
				add_content( SmartTag.new( e ).content )
			end
		}
		if( @tmpLine.length > 0 ) 
			@content.push(@tmpLine)
		end
	end
	def add_content( contents )
		contents.each { |t|
			if( @tmpLine.length + "#{t}".length > @width )
				@content.push( @tmpLine )
				@tmpLine = ""
			end
			@tmpLine += "#{t} "
		}
	end
end
class BodyItem
	attr_reader :content
	@item
	@width
	@bodyContent 
	def initialize(body, width)
		@item = body
		@width = width
		@content = Array.new()
		@item.elements.each { |e| 
			if( e.expanded_name == "w:p" )
				addContent(ParagraphItem.new(e, @width).content)
			elsif( e.expanded_name == "w:customXml" )
				# Custom XML that will hold "w:p" items.
				addContent( ParagraphItem.new(e.elements.each("w:p"), @width).content)
			elsif( e.expanded_name == "w:tbl" )
				# Parse a table
			end
		}
	end
	def addContent( newContent )
		newContent.each { |line|
			@content.push(line)
		}
	end
end


########
# Main #
########

# Bounds check on number of arguments 
if (ARGV.length > 2 )
	usage()
	Process.exit
end

# Read the arguments if they exist
ARGV.each { |arg|
	# parse for the colorized format tag.
	# -t : colorize format
	if( arg.match(/-?h+/) != nil )
		usage( )
		Process.exit
	elsif( arg.match(/-w/) != nil)
		if( ARGV.length == 2 )
			$termWidth = ARGV[1].to_i	
			ARGV.pop
			ARGV.pop
		else
			puts "Invalid Arguments"
			usage( )
			Process.exit
		end
	else 
		puts "Invalid Argument \"#{arg}\" passed."
		usage( )
		Process.exit
	end
}
#######
# Desc : Process the stdinput if it exists.
# As Per: http://blog.footle.org/2008/08/21/checking-for-stdin-inruby/
if STDIN.fcntl( Fcntl::F_GETFL, 0) == 0
	#puts "got something: #{STDIN.read}"
	$docString = getDocumentXML
else
	puts "No input from file, use \"cat <filename> | ./antiword.rb\" or use \"-h\" for help."
	usage( )
	Process.exit
end

# Create the rexml document to traverse.
docXML = REXML::Document.new($docString, {:respect_whitespace => %w{ w:t" } } )

docXML.elements.each("w:document/w:body") { |body| 
	puts BodyItem.new(body, $termWidth).content
}
puts
