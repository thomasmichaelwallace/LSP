$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Restore
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

'Restore LUSAS after an LSP script fails

call setManualRefresh(false)		'set refreshing to automatic
call suppressMessages(0)		'restore textwin messages
call unInitStatusBarProgressCtrl()	'clear any progress bars

're-call lsp menu
call fileOpen(GetSystemString("SCRIPTS") & "LSP\LSP - Menu.vbs")