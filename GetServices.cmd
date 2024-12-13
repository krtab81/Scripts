@ECHO OFF

	@SETLOCAL
	@SET DestFile=%1
	IF "%DestFile%"=="" SET DestFile=%~dp0Services.txt
	sc queryex  > %DestFile%
