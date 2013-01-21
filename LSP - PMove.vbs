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

dim from_x
dim from_y
dim from_z

dim to_x
dim to_y
dim to_z

dim vector(2)

dim point

dim trans_attr
dim transform

'select first point
set point = selection.getObjects("Points")(0)

'get current location
from_x = point.getX()
from_y = point.getY()
from_z = point.getZ()

'get new location
to_x = cdbl(inputbox("New x location", "Point re-locator", from_x))
to_y = cdbl(inputbox("New y location", "Point re-locator", from_y))
to_z = cdbl(inputbox("New z location", "Point re-locator", from_z))

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

'clean-up
call database.updateMesh()
call database.deleteAttribute(trans_attr)
