@ECHO OFF

	@SETLOCAL
	@SET DestFile=%1
	IF "%DestFile%"=="" SET DestFile=%~dp0Tasks.txt
	schtasks /Query /V > %DestFile%
