@ECHO OFF

ECHO The LUSAS Scriping Pack (LSP): Solver Daemon
ECHO Copyright (C) 2010-2013 Thomas Michael Wallace (http://www.thomasmichaelwallace.co.uk)
ECHO.
ECHO  This file is part of the LSP.
ECHO.
ECHO     The LSP is free software: you can redistribute it and/or modify
ECHO     it under the terms of the GNU General Public License as published by
ECHO     the Free Software Foundation, either version 3 of the License, or
ECHO     (at your option) any later version.
ECHO.
ECHO     The LSP is distributed in the hope that it will be useful,
ECHO     but WITHOUT ANY WARRANTY; without even the implied warranty of
ECHO     MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
ECHO     GNU General Public License for more details.
ECHO.
ECHO     You should have received a copy of the GNU General Public License
ECHO     along with The LSP.  If not, see (http://www.gnu.org/licenses/)
ECHO.

REM Waits for things to be solved

ECHO Starting daemon...
ECHO.

REM note drive
SET working_dir=%CD%
SET working_drive=%CD:~0,2%

REM remove reservations
DEL "LSP - Solverd.busy"

:WAITLOCK
	
	REM Maintain lock file
	ECHO "LOCKED" > "LSP - Solverd.lock"
	PING -n 10 localhost > NUL
	
	IF EXIST "LSP - Solverw.tmp" (
		GOTO RUNSOLVER
	) ELSE (		
		ECHO Listening for solver request... [CTRL+C to Stop]
		GOTO WAITLOCK
	)

:RUNSOLVER

REM Configure solver
ECHO.
ECHO Preparing to solve...
ECHO.

ECHO BUSY > "LSP - Solverd.busy"

@ECHO ON

SET /p data_file=<"LSP - Solverd.dta"
SET /p results_file=<"LSP - Solverd.res"
SET /p log_file=<"LSP - Solverd.lgr"
SET /p local_drive=<"LSP - Solverd.dri"
SET /p local_path=<"LSP - Solverd.pat"

REM tidy number of threads
SET /p OMP_NUM_THREADS=<"LSP - Solverd.thd"
SET OMP_NUM_THREADS=%OMP_NUM_THREADS: =%

REM set paths
%local_drive%
CD %local_path%

REM kill old files
DEL "..\..\Projects\LSP - Solverd.mys"
DEL "..\..\Projects\LSP - Solverd.dat"

REM establish local file
COPY %data_file% "..\..\Projects\LSP - Solverd.dat"

REM Run Solver
..\Lusas_S.exe "..\..\Projects\LSP - Solverd.dat"

REM Copy Back Results and Log
COPY "..\..\Projects\LSP - Solverd.mys" %results_file%
COPY "..\..\Projects\LSP - Solverd.log" %log_file%

@ECHO OFF

ECHO.
ECHO Solver finished.
ECHO.

%working_drive%
CD %working_dir%
DEL "LSP - Solverw.tmp"

REM remove reservations
DEL "LSP - Solverd.busy"

GOTO WAITLOCK