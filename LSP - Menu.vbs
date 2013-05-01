$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Menu Parser
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

'Parse the LSP menu definition and create/update the LSP menu in LUSAS, also register the LSP menu
'on start-up to make presitent.

'lusas menus
dim root			'lusas root menu
dim lsp_menu			'main lsp menu
dim sub_menu			'lsp sub menu

'menu definition file
dim filesystem			'file system access
dim text_file			'open file

'definition parser
dim item			'menu text
dim command			'menu or definition command
dim help			'help string

'script options
dim install			'check install return
dim script_path			'script access path
dim menu_path			'definition file path
dim conf_path			'path to configuration directory
dim flag_path			'appended file path
dim open_path			'after open model hook path
dim new_path			'after new model hook path
dim check_lsp_string		'expected menu name
dim install_message		'update/install prefix

'script options
script_path = GetSystemString("SCRIPTS") & "LSP\LSP - "
menu_path = script_path & "Menu.txt"
conf_path = GetSystemString("CONFIGDIR")
flag_path = conf_path & "LSP\LSP - LSP - Enabled.dat"
open_path = GetSystemString("CONFIGDIR") & "afterOpenModel.vbs"
new_path = GetSystemString("CONFIGDIR") & "afterNewModel.vbs"
check_lsp_string = "LSP"

'setup objects
set root = GetMainMenu()
set filesystem = CreateObject("Scripting.FileSystemObject")

'check to see if installation/update is needed and wanted
if root.exists(check_lsp_string) then
	install_message = "Update"
else
	install_message = "Install"
end if

'check to see if load is silent
if filesystem.fileexists(flag_path) then
	install = vbYes
else
	install = msgbox(install_message & " and register the LUSAS Scripting Pack Menu?", _
		vbQuestion + vbYesNo, "LSP Menu Installer")
end if

'provide option to skip
if install = vbYes then

	'prepare file
	set text_file = filesystem.OpenTextFile (menu_path, 1)

	'read through lsp menu defintiion file
	do until text_file.AtEndOfStream
		item = trim(text_file.ReadLine)
		command = trim(text_file.ReadLine)
		help = trim(text_file.ReadLine)

		'action definition level commands
		select case command
		
			'add menus and sub menus
			case "__main_menu__"
				if root.exists(item) then root.remove(item)								
				'try to keep menu layout as expected
				if root.exists("Window") then
					set lsp_menu = root.insertMenu("Window", item)
				else
					set lsp_menu = root.appendMenu(item)
				end if				
				'get version number
				lsp_version = help				
			case "__sub_menu__"
				set sub_menu = lsp_menu.appendMenu(item)
			
			'add seperators
			case "__main_seperator__"
				call lsp_menu.appendSeparator()				
			case "__sub_seperator__"
				call sub_menu.appendSeparator()

			'link to form editor
			case "__form_editor__"
				command = "fileOpen """ & GetSystemString("SCRIPTS") & "NewForms\OpenNewForms"""				
				call sub_menu.appendItem(item, command, help)

			'show update version and date
			case "__version__"
				command = "getLPIversion()"
				item = "LUSAS Scripting Pack (v" & lsp_version & ")"
				call sub_menu.appendItem(item, command, help)
				call sub_menu.enableItem(item, False)

			'show authorship
			case "__author__"
				item = Replace(item, "_", " ")
				command = "getLPIversion()"
				call lsp_menu.appendItem(item, command, help)
				call lsp_menu.enableItem(item, False)				

			'append menu item
			case else								
				command = "fileOpen """ & script_path & command & ".vbs""" 
				call sub_menu.appendItem(item, command, help)					
				
		end select
	loop

	'garbage collect
	text_file.close
	set text_file = nothing

	'make persistent
	if not filesystem.FileExists(flag_path) then
		set text_file = filesystem.CreateTextFile(flag_path)
		text_file.close
		set text_file = nothing
		
		'add lsp menu to open model start-up hook
		set text_file = filesystem.OpenTextFile(open_path, 8)
		text_file.writeLine ""
		text_file.writeLine "FileOpen(""" & script_path & "Menu"")"
		text_file.close
		set text_file = nothing
		
		'add lsp menu to new model start-up hook
		set text_file = filesystem.OpenTextFile(new_path, 8)
		text_file.writeLine ""
		text_file.writeLine "FileOpen(""" & script_path & "Menu"")"
		text_file.close
		set text_file = nothing		
	end if	
	
	'auto colour, if selected permenant
	call fileOpen(script_path & "Colour.vbs")
end if