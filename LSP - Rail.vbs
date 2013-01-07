$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Ballast Generator
'	Copyright (C) 2010-2012 Thomas Michael Wallace <http://www.thomasmichaelwallace.co.uk>

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

'Generates joint springs between points offset verticall from eachother (i.e. ballast).

Call RailGen()

Sub RailGen()
	'generate joint springs between points offset vertically from eachother

	'configuration defintion
	Dim group_name			'string; name of group to work on
	Dim delta						'long; z limit between low and high
	Dim mesh						'string; name of mesh to apply
	
	'array definition
	Dim point_x()				'long; x coordinate of point
	Dim point_low()			'int; id number of master point
	Dim point_high()		'int; id number of slave point
	Dim point_action()	'boolean; include point for joining

	'limit definition
	Dim point_max				'int; maximum no. of matchable points	
	Dim points					'object; total no. of points in group

	'iterator definition
	Dim point						'int; index of current working point
	Dim inspect					'object; point for inspection
	Dim bound						'int; curret upper bound

	'memory definition
	Dim x								'long; x-coordinate of current point
	Dim z								'long; z-coordinate of current point
	Dim index						'int; index of matched point
	Dim master					'object; master point
	Dim slave						'object; slave point
	Dim joint						'object; joint mesh attribute
	
	'default configuration
	group_name = cstr(inputbox("Group to connect", "Ballast Generator", "Track"))
	delta = cdbl(inputbox("Mid-height between top/bottom members", "Ballast Generator", "0.5"))
	mesh = cstr(inputbox("Joint mesh attribute name", "Ballast Generator", "JNT4"))
	
	'set extent of search
	Call selection.remove("All")
	Call selection.add("Group", group_name, "Recurse")
	point_max = selection.count("Point")
	points = selection.getObjects("Points")
	
	'reserve memory
	ReDim point_x(point_max)
	ReDim point_low(point_max)
	ReDim point_high(point_max)
	ReDim point_action(point_max)	
	bound = 0		
	
	'match points
	For point = 0 To ubound(points)

		'select point for inspeciton
		set inspect = points(point)		
		
		'test for validity of actioning
		If inspect.isMemberOfGroup(group_name) Then	
						
			'locate point
			x = inspect.getX()
			z = inspect.getZ()
			index = isIn(x, point_x, bound):

			'create match if required
			If index = -1 Then
				bound = bound + 1
				index = bound
				point_x(index) = x
			End If
			
			'switchboard for master/slave
			If z <= delta Then
				point_low(index) = point
			Else
				point_high(index) = point
			End If		
		
		End If		
	Next

	'link matched points
	For index = 1 to bound

		'select master/slave points 
		Call selection.remove("All")			
		Call selectionMem.remove("All")
		Set master = points(point_low(index))
		Set slave = points(point_high(index))
		Call selection.add(master)
		Call selection.add(slave)

		'apply join mesh
		Call assignment.setAllDefaults()
		Set joint = database.getAttribute("Point Mesh", mesh)
		Call joint.assignTo(selection, selectionMem, assignment)
	
	Next		
	
End Sub


Function isIn(find, array, bound)
	'find object in array

	Dim n	'int; iterator
	
	'iterate through array until match is found
	For n = 1 To bound
		If find = array(n) Then
			isIn = n
			Exit Function
		End If
	Next
	
	'return false if not found
	isIn = -1
	
End Function
