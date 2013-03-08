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

dim report_string	'variable look-up string
dim result		'result to report
dim type_code		'geometry type code

dim geometry		'geometry object
dim assignments		'geometry assignments
dim support		'assigned support object
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
report_string = inputbox("Write varaible", "Annotate Supports", "Ustiff")

'label styling
call db.getPen(18, pen_red, pen_green, pen_blue,1,1)
font_def = "Arial;90;Normal;NoItalic;NoUnderline;NoStrikeOut;0;"

'label setup
align = "  "
prefix = "_sup_"

'fetch every support of each selected geometry
for each geometry in selection.getObjects("All")		
	assignments = geometry.getAssignments("Support")	
	for i = lbound(assignments) to ubound(assignments)	
		set support = assignments(i).getAttribute()	
		
		'load specified report
		result = support.getValue(report_string)
		
		'attempt to find label position for different geometric types
		type_code = geometry.getTypeCode()		
		select case type_code		
			
			'points are at points
			case 1
				x = geometry.getX()
				y = geometry.getY()
				z = geometry.getZ()
			
			'lines are at mid-point
			case 2
				x = 0.5*(geometry.getStartPosition()(0) + geometry.getEndPosition()(0))
				y = 0.5*(geometry.getStartPosition()(1) + geometry.getEndPosition()(1))
				z = 0.5*(geometry.getStartPosition()(2) + geometry.getEndPosition()(2))
				
			'just 0,0,0 anything else
			case else
				x = 0
				y = 0
				z = 0		
		end select
		
		'create label as annotation
		set label = database.createAnnotationText()
		call label.setText(align & result)
		call label.setName(prefix & type_code & "_" & geometry.getID() & "_" & i)
		
		'set colour information
		call label.setColour(pen_red, pen_green, pen_blue)
		call label.setFont(font_def)
		
		'position annotation by node
		call label.setAlignTop()
		call label.setAlignLeft()
		call label.setRotation(0.0)
		call label.fixToModel()
		call label.setPosition(x, y, z)
		call label.showInAllViews()			
			
	next
next