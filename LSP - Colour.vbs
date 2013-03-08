$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Colour
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

'Apply LSP style definitions

'menu definition file
dim filesystem			'file system access
dim text_file			'open file

'definition parser
dim colour_def			'colour definition

'pen codes
dim pen_no			'pen number
dim pen_width			'default pen width
dim pen_style			'default pen style

'colour codes
dim red_code			'red byte
dim green_code			'green byte
dim blue_code			'blue byte

'colour file
dim script_path			'script access path
dim file_path			'definition file path

'script options
script_path = GetSystemString("SCRIPTS") & "LSP\LSP - "
file_path = script_path & "Colour_" & inputbox( _
	"Style name (LUSAS, Obsidian, Quartz, Sun, Moon, [custom])", _
	"Colour loader", "Obsidian") & ".txt"
pen_style = 0

'read colour definition file
set filesystem = CreateObject("Scripting.FileSystemObject")
set text_file = filesystem.OpenTextFile (file_path, 1)

'initiate user colour mode
call view.useSystemColours(false)

'show copyright message
colour_def = trim(text_file.ReadLine)
call textwin.writeLine(colour_def)

'load colour definitions
do until text_file.AtEndOfStream		
	colour_def = trim(text_file.ReadLine)
	
	'parse
	pen_no 		= cint(trim(split(colour_def,":")(0)))
	red_code 	= cint(trim(split(split(colour_def,":")(1),",")(0)))
	green_code 	= cint(trim(split(split(colour_def,":")(1),",")(1)))
	blue_code 	= cint(trim(split(split(colour_def,":")(1),",")(2)))
	pen_width 	= cint(trim(split(split(colour_def,":")(1),",")(3)))	

	'define background as pen 0
	if pen_no = 0 then
		call view.setBackgroundColour(red_code, green_code, blue_code)
	
	'set pen as defined
	else
		call database.setPen((pen_no - 1), red_code, green_code, blue_code, pen_style, pen_width)	
	
	end if
loop
	
'garbage collect
text_file.close
set text_file = nothing