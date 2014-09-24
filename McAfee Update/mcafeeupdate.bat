REM #########################################
REM #					    #
REM #	 Copyright Syscon Inc.  2008	    #
REM #	 Written By: Scott Lowell	    #
REM #					    #
REM #	 This script does a silent	    #
REM #	 update of the Virus Engine	    #
REM #	 & Dat files.			    #
REM #					    #
REM #	 Change Line 2 to the Drive Letter  #
REM #	 of the install path of Mcafee      #
REM #				    	    #
REM #	 Change Line 3 to the path of	    #
REM #	 the install directory of   	    #
REM #	 McAfee Enterprise 8.5.0i	    #
REM #					    #		
REM #	 Command Line= mcafeeupdate.bat	    #   
REM #########################################

cd\
c:
cd C:\Program Files\McAfee\VirusScan Enterprise
MCUPDATE /update /quiet
