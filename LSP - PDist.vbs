$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Point Distances
'Copyright (C) 2010-2012 Thomas Michael Wallace <http://www.thomasmichaelwallace.co.uk>

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

'Provides comprehensive information about the distance between two points.

Option Explicit

dim points		'points to operate on
dim p1			'point from
dim p2			'point to

dim dx			'delta x
dim dy			'delta y
dim dz			'delta z
dim l			'length
dim m			'gradient x:y
dim theta		'angle XOY

points = selection.getObjects("Points")

'check for sane input
if not ubound(points) = 1 then
	textwin.writeLine("Please select exactly two points.")

else
	set p1 = points(0)
	set p2 = points(1)
	
	'get differentials
	dx = p2.getX() - p1.getX()
	dy = p2.getY() - p1.getY()
	dz = p2.getZ() - p1.getZ()
	
	'calculate properties
	l = sqr(dx^2 + dy^2 + dz^2)
	if dx > 0 then
		m = dy / dx
		theta = atn(m)
	else
		m = "inf."
		theta = atn(9999)
	end if

	
	'report
	textwin.writeLine("Distance information between points " & _
		p1.getName() & " and " &  p2.getName())
	
	textwin.writeLine("| dx = " & cstr(dx))
	textwin.writeLine("| dy = " & cstr(dy))
	textwin.writeLine("| dz = " & cstr(dz))
	textwin.writeLine("| length = " & cstr(l))
	textwin.writeLine("| gradient x:y = " & cstr(m))
	textwin.writeLine("| angle XOY = " & cstr(theta))

end if
