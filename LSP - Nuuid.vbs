$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Node UUID'r
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

'Generate uuids for nodes to allow direct comparision between models

dim file_path			'uuid file path
dim filesystem		'file system conneciton
dim text_file			'report file

dim nodes					'selected nodes
dim node					'working nodes

dim node_x				'x coordinate
dim node_y				'y coordinate
dim node_z				'z coordinate

dim uuid					'uuid as x_y_z
dim node_id				'node id

'dump uuids in file path
file_path = inputbox("Label file", "UUID Labeller", "uuid") & ".csv"

'provide worse-case sort-time
nodes = selection.getObjects("Nodes")
call initStatusBarProgressCtrl("Writing selected node uuids", ubound(nodes))

'create file if required
set filesystem = CreateObject("Scripting.FileSystemObject")
if not filesystem.FileExists(file_path) then
	set text_file = filesystem.CreateTextFile(file_path)
	text_file.close
	set text_file = nothing
end if	
set text_file = filesystem.OpenTextFile(file_path, 2)

'iterate through nodes
for each node in nodes

	'determine node address
	node_x = node.getX()
	node_y = node.getY()
	node_z = node.getZ()

	'develop uuid codes
	uuid = encode_uuid(node_x, node_y, node_z)
	node_id = cstr(node.getID())
	
	'write to csv
	text_file.writeLine uuid & "," & node_id
	
	'step progress
	call statusBarProgressCtrlStep()	
next
call unInitStatusBarProgressCtrl()	

'close file
text_file.close
set text_file = nothing

function encode_uuid(node_x, node_y, node_z)
	'returns the uuid of a node from xyz coordinates

	dim str_x			'x uuid
	dim str_y			'y uuid
	dim str_z			'z uuid
	dim chr_limit	'delimiter
	
	'take next out-of-range printable character
	chr_limit = chr(33 + 100 + 2 + 1)
	
	'encode individual uuids
	str_x = encode_coord(node_x)
	str_y = encode_coord(node_y)
	str_z = encode_coord(node_z)
	
	'condense into single uuid
	encode_uuid = str_x & chr_limit & str_y & chr_limit & str_z
		
end function

function encode_coord(coord)
	'returns uuid encoded coordinate

	dim s_coord			'coordinate string
	dim uuid				'developed uuid
	
	dim codes				'coordinate codes
	dim code				'section codes
	
	dim sector			'number sector
	dim pos					'number position

	dim chr_code		'developed character code
	dim chr_offset	'zero character code
	dim chr_skip		'manage npc [delete] in mid-talbe
	dim chr_dec			'character_code of decimal
	
	'set codes
	chr_offset = 33			'start from !, ignoring [space]
	chr_skip_com = 44		'skip commas for csv
	chr_skip_del = 127	'skip non-printable [delete]
	chr_dec = chr(chr_offset + 100 + 2)
											'keep decimal point code out of range

	'encoding requires string content and decimal place
	s_coord = cstr(coord)
	if len(s_coord) = 0 then s_coord = "0.0"
	if instr(s_coord,".") = 0 then s_coord = s_coord & ".00"
	
	'expand non-destrutively into pairs
	codes = split(s_coord, ".")
	if len(codes(0)) mod 2 > 0 then codes(0) = "0" & codes(0)
	if len(codes(1)) mod 2 > 0 then codes(1) = codes(1) & "0"

	'develop uuid in number pairs
	uuid = ""
	for sector = 0 to 1	

		for pos = 0 to len(codes(sector)) / 2 - 1
	
			'convert 0-99 as ascii code avoiding non-printable sections
			chr_code = cint(mid(codes(sector), (pos * 2 + 1), 2))
			chr_code = chr_code + chr_offset
			
			'avoid non-printable in-table charactors
			if chr_code >= chr_skip_com then chr_code = chr_code + 1
			if chr_code >= chr_skip_del then chr_code = chr_code + 1			
			code = chr(chr_code)
			
			'append code and itterate
			uuid = uuid & cstr(code)
		next
		
		'use unique decimal identifier
		if sector = 0 then uuid = uuid & chr_dec
	next
	
	'return developed uuid
	encode_coord = uuid
	
end function