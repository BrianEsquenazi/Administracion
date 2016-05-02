@echo off

sqlcmd -i scriptEliminarProcedure.sql -U usuarioadmin -P usuarioadmin -S localhost\sqlserver2008
sqlcmd -i script-05-01.sql -U usuarioadmin -P usuarioadmin -S localhost\sqlserver2008
