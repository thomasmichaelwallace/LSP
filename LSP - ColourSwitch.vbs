$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Colour Switch
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

'Switch current LSP colour definition

'script options
script_path = GetSystemString("SCRIPTS") & "LSP\LSP - "
conf_path = GetSystemString("SCRIPTS") & "LSP-User\LSP - "

'fetch prefered colour scheme
file_path = script_path & "Colour_" & inputbox( _
	"Style name (LUSAS, Obsidian, Quartz, Sun, Moon, [custom])", _
	"Colour loader", "Sun") & ".txt"

'install colour definition file if exists
set filesystem = CreateObject("Scripting.FileSystemObject")
if filesystem.fileexists(file_path) then
	
	'copy file to location
	filesystem.CopyFile file_path, conf_path & "Colour.txt"
	
	'input colour swatch
	call fileOpen(script_path & "Colour.vbs")

else
	'cope with missing file
	textwin.writeLine("Colour swatch definition not found.")
end if