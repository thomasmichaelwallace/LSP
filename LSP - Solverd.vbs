$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): Solver Daemon
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

'Linker to run solver on seperate machine.

dim script_path		'path to lsp relative script directory
dim setup_path		'path to solver daemon for instllation
dim waiter_path 	'path to solver daemon waiter

dim local_path		'daemon run path
dim local_drive		'daemon run drive

dim file_base		'base path of file
dim data_file		'path to data file
dim results_file	'path to results file
dim log_file

dim base_path		'path to (local) working directory
dim lock_path		'path to lock file
dim busy_path 		'path to busy file

dim threads		'number of threads to use

dim file_system		'file system connection
dim shell		'shell object

dim batch_cmd		'command to run solver batch file

'script options
script_path = GetSystemString("SCRIPTS") & "LSP\"
setup_path = script_path & "LSP - Solverd.bat"
waiter_path = script_path & "LSP - Solverw.bat"

'solver location must be identical on daemon machine
local_path = GetSystemString("SCRIPTS")
local_drive = left(script_path, 2)

'configure file paths
file_base = db.getDBFilename()
file_base = left(file_base, len(file_base) - 4)
data_file = file_base & ".dat"
results_file = file_base & ".mys"
log_file = file_base & ".log"

'configure daemon paths
base_path = getCWD() & "\"
base_path = inputbox("Daemon path", "Solver Daemon", base_path)
if not right(base_path, 1) = "\" then base_path = base_path & "\"
lock_path = base_path & "LSP - Solverd.lock"
busy_path = base_path & "LSP - Solverd.busy"
daemon_path = base_path & "LSP - Solverd.bat"

'always attempt to install daemon- for upgrades.
set file_system = CreateObject("Scripting.FileSystemObject")
on error resume next
	call textwin.writeLine("Installing deamon to " & daemon_path)
	file_system.CopyFile setup_path, daemon_path, True
on error goto 0

'get thread mode
threads = inputbox("Thread limit", "Solver Daemon", "*")

'save database file
call db.closeAllResults()
call db.save()

'export data file
call db.exportSolver(data_file)

'build arguments
batch_cmd = """" & waiter_path & """" & " " & _
	"""" & base_path & """" & " " & _
	"""" & lock_path & """" & " " & _
	"""" & busy_path & """" & " " & _	
	"""" & data_file & """" & " " & _
	"""" & results_file & """" & " " & _
	"""" & log_file & """" & " " & _	
	local_drive & " " & _
	"""" & local_path & """" & " " & _	
	threads

'run waiter
set shell=createobject("wscript.shell")
shell.run batch_cmd, 1, True

'load results
call db.openResults(results_file)