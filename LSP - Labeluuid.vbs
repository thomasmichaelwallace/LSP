$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Label UUID
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

'Label nodes by LSP UUID from .csv file

dim uuid_path			'path to model uuid definition file
dim label_path		'path to label uuid definition file
dim filesystem		'file system access object
dim text_file			'file access object

dim uuid_def			'uuid definition to parse
dim uuid_codes()	'model uuid map to ...
dim id_codes()		'... model id map
dim label_uuid()	'labels uuid map to ...
dim label_text()	'... label text map

dim index					'current map index
dim match_id			'model/label node match index
dim id						'serach index

dim pen_red				'pen red colour component
dim pen_green			'pen green colour component
dim pen_blue			'pen blue colour component
dim font_def			'font definition string

dim align					'string prefix to position label text
dim prefix				'label name prefix

dim node					'current node
dim label					'current label

'dump uuids in file path
uuid_path = inputbox("Model uuid file", "UUID Labeller", "uuid.csv")
label_path = inputbox("Label file", "UUID Labeller", "labels.csv")

'label styling
pen_red = 136
pen_green = 138
pen_blue = 133
font_def = "Arial;90;Normal;NoItalic;NoUnderline;NoStrikeOut;0;"

'label setup
align = "  "
prefix = "_uuid_"

'prep system
set filesystem = CreateObject("Scripting.FileSystemObject")
call view.insertAnnotationLayer()

'read uuid definition file
call textwin.writeLine("Reading model uuid file...")
set text_file = filesystem.OpenTextFile (uuid_path, 1)

'load model uuid definitions
index = -1
do until text_file.AtEndOfStream		
	uuid_def = trim(text_file.ReadLine)
	index = index + 1
	
	'register [uuid, id] from uuid file
	redim preserve uuid_codes(index)
	redim preserve id_codes(index)
	uuid_codes(index) = split(uuid_def,",")(0)	
	id_codes(index) = split(uuid_def,",")(1)
	
loop
text_file.close
set text_file = nothing

'read label definition file
call textwin.writeLine("Reading label file...")
set text_file = filesystem.OpenTextFile (label_path, 1)

'load label uuid definitions
index = -1
do until text_file.AtEndOfStream		
	uuid_def = trim(text_file.ReadLine)
	index = index + 1
	
	'register [uuid, id] from uuid file
	redim preserve label_uuid(index)
	redim preserve label_text(index)
	label_uuid(index) = split(uuid_def,",")(0)	
	label_text(index) = split(uuid_def,",")(1)
	
loop
text_file.close
set text_file = nothing

'clear existing labels
call textwin.writeLine("Clearing existing...")
call visible.add("Text Annotation", "All")
call selection.add("Text Annotation", "All")
call selection.delete("Text Annotation")

'iterate through labels
call initStatusBarProgressCtrl("Labeling nodes...", index)
for index = 0 to ubound(label_uuid)

	'check for uuid match
	match_id = -1	
	for id = 0 to ubound(uuid_codes)
		if label_uuid(index) = uuid_codes(id) then
			match_id = id
			exit for
		end if
	next

	'label if matched
	if match_id > -1 then
		
		'locate node
		set node = db.getObject("Node", clng(id_codes(match_id)))		
			
		'create label as annotation
		set label = database.createAnnotationText()
		call label.setText(align & label_text(index))
		call label.setName(prefix & label_text(index))
		
		'set colour information
		call label.setColour(pen_red, pen_green, pen_blue)
		call label.setFont(font_def)
		
		'position annotation by node
		call label.setAlignTop()
		call label.setAlignLeft()
		call label.setRotation(0.0)
		call label.fixToModel()
		call label.setPosition(node.getX(), node.getY(), node.getZ())
		call label.showInAllViews()				
			
	end if
	
	'step progress
	call statusBarProgressCtrlStep()	
next
call unInitStatusBarProgressCtrl()