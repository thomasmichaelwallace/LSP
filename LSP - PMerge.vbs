$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Merge Points
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

'Force two points to merge, even if they are not geometrically coincident.

dim points		'points to operate on
dim p1			'point from
dim p2			'point to

dim x			'centre x
dim y			'centre y
dim z			'centre z

points = selection.getObjects("Points")

'check for sane input
if not ubound(points) = 1 then
	textwin.writeLine("Please select exactly two points.")
	
else
	
	'fetch pointers
	set p1 = points(0)
	set p2 = points(1)
	
	'get average position
	x = 0.5 * (p1.getX() + p2.getX())
	y = 0.5 * (p1.getY() + p2.getY())
	z = 0.5 * (p1.getZ() + p2.getZ())

	'set points to average position
	call geometryData.setAllDefaults()
	call geometryData.modifyPosition(x, y, z)
	
	'move both points
	call selection.modify(geometryData)

	'manually merge both points
	call selection.merge(geometryData)
	
end if