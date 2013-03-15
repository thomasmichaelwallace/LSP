$ENGINE=VBScript

'The LUSAS Scriping Pack (LSP): 64bit Solver
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

'Run single threaded solver; workaround for 64bit errors with complex models.


dim script_path		'path to lsp relative script directory
dim solver_path		'path to solver batch file

dim file_base		'base path of file
dim data_file		'path to data file
dim results_file	'path to results file
dim work_path		'path to (local) working directory
dim work_drive		'drive of (local) working directory

dim batch_cmd		'command to run solver batch file
dim shell		'shell object

'script options
script_path = GetSystemString("SCRIPTS") & "LSP\"
solver_path = """" & script_path & "LSP - Solver64.bat" & """"

'allow environmental options [off until batch corrected]
'env_options = inputbox("Additional Environment?", "Solver64", "OMP_SET_THREADS=1")
'env_vars = split(env_options, "=")

'save database file
call db.closeAllResults()
call db.save()

'setup file names
file_base = db.getDBFilename()
file_base = left(file_base, len(file_base) - 4)
data_file = """" & file_base & ".dat" & """"
results_file = """" & file_base & ".mys" & """"
work_path = """" & script_path & """"
work_drive = left(script_path, 2)

'export data file
call db.exportSolver(file_base & ".dat")

'build arguments
batch_cmd = solver_path & " " & work_drive & " " & _ 
	work_path & " " & data_file & " " & results_file

'on error resume next	
'	for i = 0 to ubound(env_vars)/2
'		batch_cmd = batch_cmd & " " & env_vars(i*2) & " " & env_vars(i*2+1)
'	next
'on error goto 0

'run runner
set shell=createobject("wscript.shell")
shell.run batch_cmd, 1, True

'load results
call db.openResults(file_base & ".mys")