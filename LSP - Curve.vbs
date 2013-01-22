$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Curve
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

'Curves a compound discrete load about a radii

dim compound		'compound loading object
dim discrete		'discrete loading object
dim index		'index counter for compound library
dim row			'row of position data
dim pos			'current position data
dim curved_pos		'curved position data
dim three		'definition is 3d or 2d
dim x			'x position to translate
dim y			'y position to translate

dim name		'name of compound load to apply to
dim radius		'radius to apply
dim offset		'angle to offset (+ve clockwise)

'set user options
name = Inputbox("Compound Name","Compund Curve","Cmp")
radius = cdbl(Inputbox("Curve Radius","Compound Curve","310"))
offset_angle = cdbl(Inputbox("Offset Angle (deg)","Compound Curve","0"))
offset_x = cdbl(Inputbox("Centre Offset X","Compound Curve","0"))
offset_y = cdbl(Inputbox("Centre Offset Y","Compound Curve",cstr(radius)))

'get main compound object
set compound = database.getAttribute("Discrete Compound Load", name)

'scroll through loading
for index = 0 to compound.countLoading() - 1
	set discrete = compound.getLoading(index)

	'convert patch lines to curves
	if discrete.getDiscreteLoadType() = "Patch" then LineToCurve(discrete)

	'get each application position
	for row = 0 to discrete.countRows("pos") - 1
		pos = discrete.getValue("pos", row)

		'test array dimensions to cope with LUSAS' random return values
		if IsArray(pos(0)) then
			x = pos(0)(0)
			y = pos(0)(1)
		else
			x = pos(0)
			y = pos(1)
		end if
		
		'curve and save points	
		curved_pos = CartToPara(x, y, radius, offset_angle, offset_x, offset_y)		
		call discrete.setValue("pos", curved_pos, row)
		
	next		
next


Sub LineToCurve(ByRef line)
	'convert a line patch into a curve patch

	dim row			'row counter
	dim x			'x coordinate of row
	dim y			'y coordinate of row
	dim count		'total number of defining rows (expected to be 2)
	dim rows		'new rows defining curve
	dim pdir		'original projection vector
	dim dirtype		'direction type
	dim project		'projection vector to set
	dim project_x		'projection vector x
	dim project_y		'projection vector y
	dim project_z		'projtetion vector z	
	
	'create temporary copy of data
	count = line.countRows("pos") - 1
	redim rows(count + 1, 3)
	
	'get existing coordinate points
	for row = 0 to count
		pos = line.getValue("pos", row)
		
		'test array dimensions to cope with LUSAS' random return values
		if IsArray(pos(0)) then
			x = pos(0)(0)
			y = pos(0)(1)
		else
			x = pos(0)
			y = pos(1)
		end if
	
		P = line.getValue("P", row)
				
		rows(row, 0) = cdbl(x)
		rows(row, 1) = cdbl(y)
		rows(row, 2) = cdbl(0.0)
		rows(row, 3) = cdbl(P)
	next	
	
	'create average point
	for row = 0 to ubound(rows, 2)
		rows(2,row) = 0.5 * (rows(0,row) + rows(1,row))
	next
	
	'attempt to be clever about projection
	pdir = line.getValue("PDir")
	dirtype = line.getValue("dirType")

	'load clever defaults
	project_x = array(1.0,0.0,0.0)
	project_y = array(0.0,1.0,0.0)
	project_z = array(0.0,0.0,1.0)
	
	'protect custom vectors
	if CompareArray(pdir, project_x) or CompareArray(pdir, project_y) or CompareArray(pdir, project_y) then
		project = pdir
	else
		'correct LUSAS' defaults
		if dirtype = "X" then project = project_x
		if dirtype = "Y" then project = project_y
		if dirtype = "Z" then project = project_z
	end if
		
	'create base object
	set curve = database.createLoadingDiscretePatch(line.getName())
	call curve.setDiscretePatch("line3", dirtype, project)
	call curve.setDivisions(0, 0)

	'repopulate	
	call curve.addRow(rows(0,0), rows(0,1), rows(0,2), rows(0,3))
	call curve.addRow(rows(2,0), rows(2,1), rows(2,2), rows(2,3))
	call curve.addRow(rows(1,0), rows(1,1), rows(1,2), rows(1,3))
	
End Sub


Function CompareArray(arr1, arr2)
	'Itterate through, and compare, two arrays

	Dim index

	CompareArray = False
	
	for index = 0 to ubound(arr1)
		if arr1(index) <> arr2(index) then exit function		
	next

	CompareArray = True
	
End Function


Function CartToPara(x, y, r, o, i, j)
	'Convert cartiesian coordinates to parametric circular ones
	
	'(x, y) - coordinate, r - radius, o - offset angle (deg), (i, j) - offset vector
	
	Dim s			'arc length
	Dim a			'circle centre x
	Dim b			'circle centre y
	Dim theta		'arc angle
	Dim t			'parametric angle	
	Dim c			'parametric radius
	Dim x_p			'parametric x
	Dim y_p			'parametric y
	Dim para(2)		'parametric array
	Dim coord(2)		'coordinate array
	Dim blank(2)		'blank array
	Dim pi			'the tasty constant
	Dim delta		'offset angle
	
	'assign pi, which is inexplicably outside of the vb scope
	pi = 4.0 * atn(1.0)
	
	'define parametric arc
	s = x
	c = r + y
	
	'convert angle to parametric angle
	theta = s/r
	t = 0.5 * pi - theta
	
	'set offset
	if o <> 0.0 then
		delta = (2.0 * pi * o) / 360.0
	else
		delta = o
	end if
	
	'apply offset
	t = t + delta
	
	'find circle centre
	a = 0.0
	b = -r
	
	'convert to parametric coordinates
	x_p = a + c * cos(t)
	y_p = b + c * sin(t)
	
	'apply offset vector
	x_p = x_p + i
	y_p = y_p + j
	
	'assign parametric points
	para(0) = cdbl(x_p)
	para(1) = cdbl(y_p)
	para(2) = cdbl(0.0)

	'create dummy object for third dimension definition
	blank(0) = cdbl(0.0)
	blank(1) = cdbl(0.0)
	blank(2) = cdbl(0.0)	
	
	'package for three dimensions
	coord(0) = para
	coord(1) = blank
	coord(2) = blank	
	CartToPara = coord

End Function