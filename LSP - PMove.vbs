$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Point Re-Locator
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

'Precisely re-loacate point to position.

dim from_x	'initial position x
dim from_y	'initial position y
dim from_z	'initial positions z

dim to_x	'final position x
dim to_y	'final position y
dim to_z	'final position z

dim read_x	'input position x
dim read_y	'input position y
dim read_z	'input position z

dim vector(2)	'movement vector

dim point	'point to move

dim trans_attr	'transformation attribute
dim transform	'transformation utility

'get new location
read_x = inputbox("New x location (use @ for current)", "Point re-locator", "@")
read_y = inputbox("New y location (use @ for current)", "Point re-locator", "@")
read_z = inputbox("New z location (use @ for current)", "Point re-locator", "@")

'select first point
for each point in selection.getObjects("Points")

	'get current location
	from_x = point.getX()
	from_y = point.getY()
	from_z = point.getZ()

	'run substitutions
	if read_x = "@" then
		to_x = from_x
	else
		to_x = cdbl(read_x)
	end if
	if read_y = "@" then
		to_y = from_y
	else
		to_y = cdbl(read_y)
	end if
	if read_z = "@" then
		to_z = from_z
	else
		to_z = cdbl(read_z)
	end if
	'create vector
	vector(0) = to_x - from_x
	vector(1) = to_y - from_y
	vector(2) = to_z - from_z
	
	'create temporary translation
	set trans_attr = database.createTranslationTransAttr("__LSP_PMove", vector)
	set transform = database.getTransformation("__LSP_PMove")

	'prepare for vanilla movement
	call geometryData.setAllDefaults()
	call geometryData.setTransformation(transform)

	'run move
	call selection.move(geometryData)

next
	
'clean-up
call database.updateMesh()
call database.deleteAttribute(trans_attr)
