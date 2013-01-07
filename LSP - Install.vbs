$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Web-Installer
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

'Download and install/update current LSP.

'github url
github_url = "https://github.com/thomasmichaelwallace/LSP/archive/master.zip"

'lsp to be installed locally to lusas drive
lsp_path = GetSystemString("SCRIPTS") & "LSP"
zip_file = lsp_path & "-master.zip"

'create file system connection
set filesystem = CreateObject("Scripting.FileSystemObject")

'create and make an xml http request
set xml_http = CreateObject("MSXML2.XMLHTTP")
xml_http.open "GET", github_url, false
xml_http.send()

'check request response
If xml_http.Status = 200 Then
	textwin.writeLine("Downloading latest version of the LSP from GitHub...")

	'open binary stream
	Set ado_stream = CreateObject("ADODB.Stream")
	ado_stream.Open
	ado_stream.Type = 1 

	'write binary stream
	ado_stream.Write xml_http.ResponseBody
	ado_stream.Position = 0

	'override file if exists
	If filesystem.Fileexists(zip_file) Then filesystem.DeleteFile zip_file
	ado_stream.SaveToFile zip_file
	
	ado_stream.Close
	Set ado_stream = Nothing
Else

	textwin.writeLine("Could not connect to GitHub, Error: " & cstr(xml_http.Status))
End if

'close stream
Set xml_http = Nothing

'create lsp folder, and remove existing if required
if filesystem.FolderExists(lsp_path) then filesystem.DeleteFolder lsp_path, true
filesystem.CreateFolder(lsp_path)

'extract contents of zip file
set unzip = CreateObject("Shell.Application")
set zipped_files = unzip.NameSpace(zip_file).items
unzip.NameSpace(lsp_path).CopyHere(zipped_files)

'grabage collect
Set filesystem = Nothing
Set unzip = Nothing

'run menu installer
call fileOpen(lsp_path & "\LSP - Menu.vbs")