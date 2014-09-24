@echo off
rem ##########################################
rem # Copyright 2006 by N-able Technologies  #
rem # All Rights Reserved                    #
rem # May not be reproduced or redistributed #
rem # Without written consent from N-able    #
rem # Technologies			     #
rem # www.n-able.com			     #
rem ##########################################
rem #
rem # Defrag.bat
rem #
rem # This batch file will execute a command line  
rem # defrag of the specified drive.
rem #
rem # It must be run on the device who's service
rem # you wish to control.
rem #
rem # This script is suitable for remote execution
rem # as no UI interaction is required.
rem # 
rem # note the command line parameters.  
rem # Parameter 1: Drive to defrag.
rem # ALL PARAMETERS ARE REQUIRED.
rem #
rem # ex:  defrag.bat c:
rem
defrag %1