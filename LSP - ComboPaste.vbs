$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Combo Paste
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

'Pastes load combination defintiions described by LSP Combo Copy command into model.

'collections
dim values			'array of loaded variables

'objects
dim loading			'load being copied
dim value			'value of variable to be copied

'counters
dim row				'current internal loading row being pasted
dim rows			'total number of internal loading rows

'storage
dim paste_string		'complete data string from file
dim loading_type		'loading type being copied
dim loading_name		'name of loading being copied

'filesytem objects
dim filesystem			'file system access
dim text_file			'text file to save copying data

'options
dim file_path			'file path to save copied data

'script level defaults and options
file_path = GetSystemString("CONFIGDIR") & "LSP\LSP - ComboCopy.txt"

'prepare file
set filesystem = CreateObject("Scripting.FileSystemObject")
if not filesystem.FileExists(file_path) then
	set text_file = filesystem.CreateTextFile(file_path)
	text_file.close
	set text_file = nothing
end if 
set text_file = filesystem.OpenTextFile (file_path, 1)

'input copying file
do until text_file.AtEndOfStream

	'get copying line data
	paste_string = text_file.ReadLine
	values = split(paste_string, ";")	
	loadset_type = values(0)
	loadset_name = values(1)
	
	'check for existing
	if db.existsLoadset(loadset_name) or left(loadset_type, 2) = "|-" then
		if not left(loadset_type, 2) = "|-" then _
			call textwin.writeLine("Skipping existing: " & loadset_name)
	
	else		
		call textwin.writeLine("Pasting loadset: " & loadset_name)
	
		select case loadset_type
		
			case 2
				set loadset = db.createCombinationBasic(loadset_name)				
				
				paste_string = text_file.ReadLine
				paste_string = mid(paste_string, 6)
				factors = split(paste_string, ";")
				
				paste_string = text_file.ReadLine
				paste_string = mid(paste_string, 6)
				loadcases = split(paste_string, ";")				
			
				complete = true
				for each loadcase in loadcases
					if not db.existsLoadset(loadcase) then
						complete = false
						db.deleteLoadset(loadset)
						textwin.writeLine(loadcase & " missing, aborted.")
						exit for
					end if
				next
				
				if complete then call loadset.addEntries(factors, loadcases)
			
			case else
				if not isNumeric(loadset_name) then _
					call textwin.writeLine("Loading type unhandled.")				
				
		end select
	end if
loop

'clean up
text_file.close