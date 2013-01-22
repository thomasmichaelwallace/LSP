$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Resize Curve
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

'Resize a patch load curved using the LSP Curve command.

'interpetors
dim left			'left-most point
dim row				'data row

'objects
dim discrete			'patch load to alter
dim pos				'LUSAS point position

'circle definition
dim radius			'circle radius
dim ab				'circle centre x;y

'coordinate relations
dim coords(2,1)			'coordinates of patch load curve
dim chords(2)			'chord lengths between definition

'parametric angles
dim t_abc(2)			'original parametric distances
dim t_efg(2)			'altered parametric distances
dim t(2)			'transformed parametric distances

'arcs
dim S				'original arc length
dim dS				'arc length change
dim S_B				'arc length to mid-point

'adjusters
dim i				'ratio between mid-point and total arc length
dim alpha			'angle increase from fixed point to alter length

'arc angles
dim theta			'original arc angle
dim dtheta			'arc angle change
dim theta_B			'arc angle to mid-point

'user options
dim name			'name of patch load to alter
dim L				'length to set patch load
dim fixed			'fixed side

'pick up attribute from tree view
name = "__Nothing__"
if treeSelection is nothing then name = Inputbox("Patch Name","Resize Curve","pch")
on error resume next
	if treeSelection.getSubType() = "Discrete Patch Load" then name = treeSelection.getName()
on error goto 0

if name = "__Nothing__" then
	msgbox "A discrete patch load must be selected in the tree view.", vbError, "Resize Curve"
	call stopScript()
end if	

L = cdbl(Inputbox("New Length","Resize Curve","40"))
fixed = Inputbox("Fixed Side (L/R)","Resize Curve","L")

'get patchload compound object
set discrete = database.getAttribute("Discrete Patch Load", name)

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
	
	'load curve points
	coords(row, 0) = x
	coords(row, 1) = y	
	
next		

'get chord lengths
chords(0) = ChordLength(coords(0,0), coords(0,1), coords(1,0), coords(1,1))
chords(1) = ChordLength(coords(1,0), coords(1,1), coords(2,0), coords(2,1))
chords(2) = ChordLength(coords(0,0), coords(0,1), coords(2,0), coords(2,1))

'set radius and circle centre
radius = TriRadius(chords(0), chords(1), chords(2))
ab = Centre(radius, coords(1,0), coords(1,1), coords(0,0), coords(0,1), coords(2,0), coords(2,1))

'fetch parametric angles
t_abc(0) = Parametric(radius, ab(0), ab(1), coords(0,0), coords(0,1))
t_abc(1) = Parametric(radius, ab(0), ab(1), coords(1,0), coords(1,1))
t_abc(2) = Parametric(radius, ab(0), ab(1), coords(2,0), coords(2,1))

'calculate differential
theta = t_abc(2) - t_abc(0)
S = theta * radius
dS = L - S
dtheta = dS / radius

'select modes
if coords(0,0) <= coords(2,0) then
	left = "A"
else
	left = "C"
end if

'preserve radios
if left = "A" then
	theta_B = t_abc(0) - t_abc(1)
else
	theta_B = t_abc(2) - t_abc(1)
end if
S_B = theta_B - radius
i = S / S_B

'define extension
alpha = theta + dtheta
if fixed = "L" then alpha = -alpha

'create extension matrix
if left = "A" then
	if fixed = "L" then
		t_efg(0) = t_abc(0)
	else
		t_efg(0) = t_abc(2)
	end if
else
	if fixed = "L" then
		t_efg(0) = t_abc(2)
	else
		t_efg(0) = t_abc(0)
	end if
end if
t_efg(2) = t_efg(0) + alpha
t_efg(1) = (t_efg(2) - t_efg(0)) * i + t_efg(0)

'preserve initial layout
if fixed = "L" then
	if left = "A" then
		t(0) = t_efg(0)
		t(2) = t_efg(2)
	else
		t(0) = t_efg(2)
		t(2) = t_efg(0)
	end if
else
	if left = "A" then
		t(0) = t_efg(2)
		t(2) = t_efg(0)		
	else
		t(0) = t_efg(0)
		t(2) = t_efg(2)			
	end if
end if
t(1) = t_efg(1)

'lusas requires curve at mid-point
t(1) = (t(0) + t(2))/2

'convert and save new points
for row = 0 to discrete.countRows("pos") - 1
	pos = ParaToCart(radius, ab(0), ab(1), t(row))
	call discrete.setValue("pos", pos, row)
next


Function ChordLength(x1, y1, x2, y2)
	'returns the length of a choord	
	ChordLength = sqr((x1-x2)^2 + (y1-y2)^2)
End Function

Function TriRadius(a, b, c)
	'return the radius of a triangle
	TriRadius = (a*b*c)/sqr((a+b+c)*(-a+b+c)*(a-b+c)*(a+b-c))
End Function
	
Function Centre(R, x, y, x1, y1, x2, y2)
	'return the centre of a circle from a point and a chord
	
	dim mx			'chord length in x direction
	dim my			'chord length in y direction
	dim h			'chord length
	dim mh			'ratio of chord length to radius
	dim a			'centre x
	dim b			'centre y
	
	'global coordinate lengths
	mx = x2 - x1
	my = y2 - y1
	
	'local coordinate lengths
	h = sqr((x2-x1)^2 + (y2-y1)^2)
	mh = R/h
	
	'similar triangles with perpendicular rotation (assuming north circle)
	a = x + mh * my
	b = y - mh * mx
	
	'return results in array
	Centre = array(a, b)

End Function	

Function Parametric(R, a, b, x, y)
	'return the parametric angle of a point on a circle
		
	dim tx			'parametric angle from x	
	dim ty			'parametric angle from y
	
	'establish posibilities
	tx = acs((x-a)/R)
	ty = asn((y-b)/R)
	
	'select from trig-cricle
	if ty > 0 then
		Parametric = tx
	else
		Parametric = 2*pi() - tx
	end if
	
End Function

Function ParaToCart(R, a, b, t)
	'return a LUSAS position from parametric coordinates
	
	dim x			'cartiesian x
	dim y			'cartiesian y
	dim para(2)		'cartiesian collection
	dim blank(2)		'blank vector
	dim coord(2)		'LUSAS vector
	
	'convert to coordinates
	x = a + R*cos(t)
	y = b + R*sin(t)

	'assign parametric points
	para(0) = cdbl(x)
	para(1) = cdbl(y)
	para(2) = cdbl(0.0)

	'create dummy object for third dimension definition
	blank(0) = cdbl(0.0)
	blank(1) = cdbl(0.0)
	blank(2) = cdbl(0.0)	
	
	'package for three dimensions
	coord(0) = para
	coord(1) = blank
	coord(2) = blank	
	ParaToCart = coord

End Function

'missing vbscript trigonomic functions
Function pi()
    pi = 4 * atn(1)
End Function
Function asn(theta)
		asn = 2 * atn(theta/(1 + sqr(1-(theta^2))))
End Function
Function acs(theta)
    acs = pi/2 - asn(theta)
End Function
