$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Loadcase Remover
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

'Mass delete range of loadsets.

dim id_from		'loadcase id to start deleteing at
dim id_to		'loadcase id to finish deleteing at
dim lcid		'loadcase id to delete
dim repeat_mode		'repeat switch

'be tolerant of missing loadcases
on error resume next

repeat_mode = vbYes

do while repeat_mode = vbYes

	'configure modeller
	setManualRefresh(true)

	'get deletion range
	id_from = cint(Inputbox("Delete from loadcase id", "Mass Delete", ""))
	id_to = cint(Inputbox("Delete to loadcase id", "Mass Delete", ""))
	
	'remove range
	for lcid = id_from to id_to
		call database.deleteLoadset(cstr(lcid))
	next

	repeat_mode = msgbox("Delete another range of loadcases?", vbYesNo, "Mass Delete")
	
	'configure modeller
	setManualRefresh(false)

loop