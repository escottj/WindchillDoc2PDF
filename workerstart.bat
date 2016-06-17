@echo off
set DOC_WORKER_INSTALL=E:\ptc\DocWorker
set JAVA_HOME=C:\Program Files\Java\jre1.8.0_91
set PATH=%PATH%;%JAVA_HOME%\bin
set CLASSPATH=%DOC_WORKER_INSTALL%\codebase\wvs.jar;%CLASSPATH%
set PS_SCRIPT=%DOC_WORKER_INSTALL%\bin\WordPDF.ps1
set DEBUG="-D"
set PORT="5600"
set HOST="PDMLINK"
set TYPE="OFFICE"
set DIR="%DOC_WORKER_INSTALL%\tmp"
set LOG="worker_"
set CMD="%DOC_WORKER_INSTALL%\bin\docworker.bat"
java com.ptc.wvs.server.cadagent.GenericWorker %DEBUG% -PORT %PORT% -HOST %HOST% -TYPE %TYPE% -CMD %CMD% -DIR %DIR% -LOG %LOG%
