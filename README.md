# -LogParserWrapper_Netlogon
Convert netlogon logs to Excel for more insight
# Readme..
	# This script will generate Status/User/Machine of SAMLogon activties summary Excel sheet from netlogon.log using LogParser and Excel.
	#		1. To enable Netlogon debug logging: run NLtest. Logging will start right after NLtest for later OS, only restart Netlogon service if debug info is not present. 
	#					Nltest /DBFlag:2080FFFF
	#		2. Output in: %windir%\debug\netlogon.log & netlogon.bak
	#		3. To stop netlogon debug logging: 
	#					Nltest /DBFlag:0
	#		4. No need to delete Netlogon.* since OS will still log esseitanl netlogon info.
	#		5. More info https://docs.microsoft.com/en-us/troubleshoot/windows-client/windows-security/enable-debug-logging-netlogon-service
	#
	# LogParserWrapper_Netlogon.ps1 v0.6 11/2
	# 	Steps:
	#   	1. Install LogParser 2.2 from https://www.microsoft.com/en-us/download/details.aspx?id=24659
	#    			Note: More about LogParser2.2 https://docs.microsoft.com/en-us/previous-versions/windows/it-pro/windows-xp/bb878032(v=technet.10)?redirectedfrom=MSDN
	#   	2. Copy Netlogon.log & Netlogon.bak from traget's %windir%\debug directory to same directory as this script.
	#     		Note1: Script will rename Netlogon.bak to Netlogon_bak.log.
	#					Note2: Afterward, LogParser will process *.log(s) in script directory in (3).
	#   	3. Run script
	# 
