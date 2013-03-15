@ECHO OFF

ECHO The LUSAS Scriping Pack (LSP): 64bit Solver
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

REM Run single threaded solver; workaround for 64bit errors with complex models.

REM set to single threaded mode
@ECHO ON
SET OMP_NUM_THREADS=1

REM set paths
%1
CD %2

REM kill old files
DEL "..\..\..\Projects\LSP - Solver64.mys"
DEL "..\..\..\Projects\LSP - Solver64.dat"

REM establish local file
COPY %3 "..\..\..\Projects\LSP - Solver64.dat"

REM Run Solver
..\..\Lusas_S.exe "..\..\..\Projects\LSP - Solver64.dat"

REM Copy Back Results
COPY "..\..\..\Projects\LSP - Solver64.mys" %4