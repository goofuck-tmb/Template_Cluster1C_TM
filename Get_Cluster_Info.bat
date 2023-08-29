@echo off

"%windir%\System32\cscript.exe" /nologo C:\zabbix\scripts\Cluster1C\Get_Cluster_Info.vbs
powershell -executionpolicy RemoteSigned -WindowStyle Hidden -file C:\zabbix\scripts\Cluster1C\getLic.ps1