$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Annotate Supports
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

'Annotate objects with support conditions

dim vector		'coordinate vector to report
dim position		'result to report

dim point		'geometry object
dim label		'annotation object

dim font_def		'label font definition
dim align		'label alginment prefix
dim prefix		'label name prefix

dim pen_red		'pen red content
dim pen_green		'pen green content
dim pen_blue		'pen blue content

dim x			'label x position
dim y			'label y position
dim z			'label z position

'allow user to specify variable to report
vector = ucase(cstr(inputbox("Write coordinate", "Annotate Points", "z")))

'label styling
call db.getPen(18, pen_red, pen_green, pen_blue,1,1)
font_def = "Arial;90;Normal;NoItalic;NoUnderline;NoStrikeOut;0;"

'label setup
align = "  "
prefix = "_pnt_"

'fetch every support of each selected point
for each point in selection.getObjects("Points")
			
	'load specified position
	select case vector
	case "X"
		position = cstr(point.getX())
	case "Y"
		position = cstr(point.getY())
	case "Z"
		position = cstr(point.getZ())
	case else
		position = ""
	end select

	'create label as annotation
	set label = database.createAnnotationText()
	call label.setText(align & position)
	call label.setName(prefix & type_code & "_" & point.getID() & "_" & vector)
	
	'set colour information
	call label.setColour(pen_red, pen_green, pen_blue)
	call label.setFont(font_def)
	
	'position annotation by node
	call label.setAlignTop()
	call label.setAlignLeft()
	call label.setRotation(0.0)
	call label.fixToModel()
	call label.setPosition(point.getX(), point.getY(), point.getZ())
	call label.showInAllViews()			
				
next