$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Import Attributes from File
'Copyright (C) 2010-2013 Thomas Michael Wallace <http://www.thomasmichaelwallace.co.uk>

' This file is part of the LSP.

'    The LSP is free software: you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation, either version 3 of the License, or
'    (at your option) any later version.

'    The LSP is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.

'    You should have received a copy of the GNU General Public License
'    along with The LSP.  If not, see <http://www.gnu.org/licenses/>.

'Reads in an attribute file from an LSP prepared header.

dim attr_type		'attribute type
dim attr_id		'attribute unique id
dim template		'attribute to build header

dim proper_type		'proper lusas type
dim proper_subtype	'proper lusas sub-type

dim proper_line		'formed line for proper type header
dim name_line		'formed line of value names
dim type_line		'formed line of value types
dim desc_line		'formed line of value descriptions
dim value_line		'formed line of example values

dim names		'value names
dim i			'counter for current name

dim name		'current value name
dim vbtype		'current value type
dim desc		'current value description
dim value		'current value example value

dim file_path		'header file path
dim filesystem		'file system connection
dim text_file		'header file

'open file
file_path = "attr_import.csv"
set filesystem = CreateObject("Scripting.FileSystemObject")
set text_file = filesystem.OpenTextFile(file_path, 1)

'fetch proper identification
proper_line = text_file.readLine()
proper_type = split(proper_line, ",")(0)
proper_subtype = split(proper_line, ",")(1)

'fetch import names
name_line = text_file.readLine() 
names = split(name_line, ",")

'skip header lines
text_file.readLine()	'type
text_file.readLine()	'descriptions
text_file.readLine()	'example

'read in values
do until text_file.AtEndOfStream
	value_line = text_file.readLine()
	values = split(value_line, ",")

	'implement custom db.existAttribute due to "type name" mismatch.
	on error resume next
		set attr = nothing
		set attr = db.getAttribute(proper_type, values(0))	
	on error goto 0
	
	if attr is nothing then
		'create attributes that don't (needs explicit typing)
	
		'attribute creation switch board
		select case proper_type
			
			case "Structural Support"
				set attr = db.createSupportStructural(values(0))		
			
			case else
				'cope with limited switchboard approach
				call textwin.writeLine(proper_type & "/" & proper_subtype & " is not currently supported.")
				set attr = nothing
		end select
	end if
	
	'only proceed if attribute type is supported
	if not attr is nothing then
	
		'itterate and load values
		for i = 1 to ubound(names)
			name = names(i)
			value = values(i)
		
			'values as * will be skipped; allows partial edit.
			if value <> "*" then
				'set values
				call attr.setValue(name, value)					
			end if

		next	
	end if
loop