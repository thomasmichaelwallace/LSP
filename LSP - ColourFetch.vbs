$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): ColourFetch
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

'Extract LSP compliant colour schemes

Dim index		'pen number
Dim red 		'red component
Dim green 		'green component
Dim blue 		'blue component
Dim style 		'pen style
Dim width 		'pen width
Dim swatchString	'string describing swatch

'initalise variables to pass to LPI
index = 1
red = 1
green = 1
blue = 1
style = 1
width = 1

'add header to seperate text 
Call textwin.writeLine(" ")
Call textwin.writeLine("LSP: Fetch Colour Scheme")
Call textwin.writeLine("------------------------")
Call textwin.writeLine(" ")

'colour scheme information
Call textwin.writeLine("Current Scheme, Copyright (C) 2013 The User")

'background cannot be fetched, so request it instead
swatchString = InputBox("Manually set background colour: ", _
	"Fetch Colour Scheme", "0: 255, 255, 255, 1")
Call textwin.writeLine(swatchString)

'itterate through pens and append details to text window
For index = 0 to 19
	Call db.getPen(index, red, green, blue, style, width)
	
	'LSP file format is index: red, green, blue, width (style not included)
	swatchString = _ 
		(index + 1) & ": " & _
		red & ", " & _
		green & ", " & _
		blue & ", " & _
		width		

	Call textwin.writeline(swatchString)		
Next