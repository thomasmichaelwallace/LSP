$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Beam Diagram
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

'Generate force/moment diagrams for selected beam.

dim mode		'mode string for results
dim entities		'entity array for results
dim entitiy		'current entity of results

dim points		'defining points
dim start_point		'start point
dim end_point		'end point

dim phi			'view rotation about x-axis
dim theta		'view rotation about y-axis
dim psi			'view rotation about z-axis

dim dx			'delta xx
dim dy			'delta yy

dim pi			'the tastiest constant

dim matrix(3,3)		'lusas view rotation matrix
dim old_matrix		'current lusas view rotation matrix

'remember previous layout
set old_visible = newObjectSet()
call old_visible.add(visible)
call suppressMessages(1)
call view.saveView("LSP_Diagrams")
call suppressMessages(0)

'script options
mode = "Force/Moment - Thick 3D Beam"
entities = array("Fx", "Fy", "Fz", "Mx", "My", "Mz")

'extract start/end points and lines from selection
lines = selection.getObjects("Lines")
set start_point = lines(0).getLOFs()(0)
set end_point = lines(ubound(lines)).getLOFs()(1)

'get user options
label = inputbox("Diagram Label", "Quick Diagram", cstr(start_point.getID()))

'get geometry
dx = end_point.getX() - start_point.getX()
dy = end_point.getY() - start_point.getY()
dz = end_point.getZ() - start_point.getZ()

'assign angles
pi = 4.0 * atn(1.0)
phi = pi * 0.5
if dz > 0 then
	theta = atn(dx/dz) + pi * 0.5
else
	theta = 0
end if
if dx = 0 then
	psi = pi * 0.5
else
	psi = atn(dy/dx)
end if

'develop matrix
matrix(0,0) = cos(theta) * cos(psi)
matrix(0,1) = -cos(phi) * sin(psi) + sin(phi) * sin(theta) * cos(psi)
matrix(0,2) = sin(phi) * sin(psi) + cos(phi) * sin(theta) * cos(psi)
matrix(0,3) = 0.0

matrix(1,0) = cos(theta) * sin(psi)
matrix(1,1) = cos(phi) * cos(psi) + sin(phi) * sin(theta) * sin(psi)
matrix(1,2) = -sin(phi) * cos(psi) + cos(phi) * sin(theta) * sin(psi)
matrix(1,3) = 0.0

matrix(2,0) = -sin(theta)
matrix(2,1) = sin(phi) * cos(theta)
matrix(2,2) = cos(phi) * cos(theta)
matrix(2,3) = 0.0

matrix(3,0) = 0.0
matrix(3,1) = 0.0
matrix(3,2) = 0.0
matrix(3,3) = 1.0

'set view
call view.setRotationMatrix( _
	cdbl(matrix(0,0)), cdbl(matrix(0,1)), cdbl(matrix(0,2)), cdbl(matrix(0,3)), _
	cdbl(matrix(1,0)), cdbl(matrix(1,1)), cdbl(matrix(1,2)), cdbl(matrix(1,3)), _
	cdbl(matrix(2,0)), cdbl(matrix(2,1)), cdbl(matrix(2,2)), cdbl(matrix(2,3)), _
	cdbl(matrix(3,0)), cdbl(matrix(3,1)), cdbl(matrix(3,2)), cdbl(matrix(3,3)) _
											)

'setup screen											
call visible.keep(selection)
call selection.remove("All")
call view.showViewSummary(true)
call view.setScaledToFit(true)
call view.setScaledToFit(false)

'create diagrams layer
call view.insertDiagramsLayer()
call view.diagrams.showInScreenPlane(true)
call view.diagrams.setShowLabels(true, 3, "Arial;90;Normal;NoItalic;NoUnderline;NoStrikeOut;-50;", true)
call view.diagrams.setLabelDecimalPlaces(3)

'setup layers for each entity
for each entity in entities	
	call view.diagrams.setResults(mode, entity)
	
	'reset view	
	call view.update()
	
	'create image
	call getCurrentView().savePicture("Drgs - " & cstr(label) & _
		" (" & entity & ").bmp", "Bitmap")
	
next

'reset back to normal
visible.add(old_visible)
call view.loadView("LSP_Diagrams", true, true, true, true, true)