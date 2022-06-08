@echo off 
if not "%OS%"=="Windows_NT" exit
title WindosActive

@REM echo before：%cd%
@REM echo to: %~dp0
cd /D %~dp0
@REM echo after：%cd%
node index.js --export