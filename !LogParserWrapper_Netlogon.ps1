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
function CommonWorkBookTasks($In) { # [WorkBook,Column,Name] Set filter, split panel, lock on row 1, auto column width. Set # column format >> save and delete CSV
		$iSheet = $Excel.Workbooks[$In[0]].Worksheets[1]
			$iSheet.Range("A1").AutoFilter() | Out-Null
			$iSheet.Application.ActiveWindow.SplitRow=1  
			$iSheet.Columns.AutoFit() | Out-Null
			$iSheet.Application.ActiveWindow.FreezePanes = $true
			$iSheet.Columns.Item($In[1]).numberformat = "###,###,###,###,###"
			$iSheet.Name = ($In[2])
			$iSheet.Cells.Item(1,5)=$In[3]
		$Excel.Workbooks[$In[0]].SaveAs($ScriptPath+'\'+$TimeStamp+'_'+$In[2],51)
		$iCSV = $ScriptPath+'\'+$TimeStamp+'_'+$In[2]+'.csv'
		Remove-Item $iCSV
}
#------Main---------------------------------
	$ScriptPath = Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path
	Get-ChildItem -Path $ScriptDirectory -Filter '*.bak' | Rename-Item -NewName {$_.name -replace '\.bak$', "_bak.log"} -ErrorAction Stop | Out-Null
		$InFiles = $ScriptPath+'\*.log'
		$InputFormat = New-Object -ComObject MSUtil.LogQuery.TextLineInputFormat
		$TimeStamp = "{0:yyyy-MM-dd_hh-mm-ss_tt}" -f (Get-Date)
		$LPQuery = New-Object -ComObject MSUtil.LogQuery
		$OutputFormat = New-Object -ComObject MSUtil.LogQuery.CSVOutputFormat
#--SamLogon-Machine_
	$OutTitle1 = 'SamLogon-Machine'
	$OutFile1 = $ScriptPath+'\'+$TimeStamp+'_'+$OutTitle1+'.csv'
	$Query = @"
		SELECT 
			CASE EXTRACT_SUFFIX(TEXT,0,'Returns ')
				WHEN '0XC000005E' THEN '5E_NO_LOGON_SERVERS' 				WHEN '0xC0000064' THEN '64_NO_SUCH_USER'
				WHEN '0xC000006A' THEN '6A_STATUS_WRONG_PASSWORD'		WHEN '0XC000006D' THEN '6D_LOGON_FAILURE'
				WHEN '0XC000006E' THEN '6E_ACCOUNT_RESTRICTION'			WHEN '0xC000006F' THEN '6F_INVALID_LOGON_HOURS'
				WHEN '0xC0000070' THEN '70_INVALID_WORKSTATION'			WHEN '0xC0000071' THEN '71_PASSWORD_EXPIRED'
				WHEN '0xC0000072' THEN '72_ACCOUNT_DISABLED'				WHEN '0XC00000DC' THEN 'DC_INVALID_SERVER_STATE'
				WHEN '0XC0000133' THEN '133_TIME_DIFFERENCE_AT_DC'	WHEN '0XC000015B' THEN '15B_LOGON_TYPE_NOT_GRANTED'
				WHEN '0xC0000193' THEN '193_ACCOUNT_EXPIRED'				WHEN '0xC0000234' THEN '234_ACCOUNT_LOCKED_OUT'
				WHEN '0x0' THEN 'OK' END AS Status, 
			TO_UPPERCASE (extract_prefix(extract_suffix(TEXT, 0, 'logon of '), 0, 'from ')) as User, 
			TO_UPPERCASE (extract_prefix(extract_suffix(TEXT, 0, 'from '), 0, 'Returns')) as MachineName, 
			COUNT(*) as Total
		INTO $OutFile1
		FROM $InFiles
		WHERE 
			INDEX_OF(TO_UPPERCASE (TEXT),'SAMLOGON') >0 AND INDEX_OF(TO_UPPERCASE(TEXT),'RETURNS') >0 AND NOT INDEX_OF(TO_UPPERCASE(TEXT),'KERBEROS') >0 
		GROUP BY 
			Status, User, MachineName ORDER BY Total DESC
"@
	Write-Progress -Activity "Generating $OutTitle1 report" -PercentComplete (30)
	$LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)| Out-Null
#--SamLogon-Domain_
	$OutTitle2 = 'SamLogon-Domain'
	$OutFile2 = $ScriptPath+'\'+$TimeStamp+'_'+$OutTitle2+'.csv'
	$Query = @"
		SELECT
			CASE EXTRACT_SUFFIX(TEXT,0,'Returns ')
				WHEN '0XC000005E' THEN '5E_NO_LOGON_SERVERS' 			WHEN '0xC0000064' THEN '64_NO_SUCH_USER'
				WHEN '0xC000006A' THEN '6A_STATUS_WRONG_PASSWORD'		WHEN '0XC000006D' THEN '6D_LOGON_FAILURE'
				WHEN '0XC000006E' THEN '6E_ACCOUNT_RESTRICTION'			WHEN '0xC000006F' THEN '6F_INVALID_LOGON_HOURS'
				WHEN '0xC0000070' THEN '70_INVALID_WORKSTATION'			WHEN '0xC0000071' THEN '71_PASSWORD_EXPIRED'
				WHEN '0xC0000072' THEN '72_ACCOUNT_DISABLED'			WHEN '0XC00000DC' THEN 'DC_INVALID_SERVER_STATE'
				WHEN '0XC0000133' THEN '133_TIME_DIFFERENCE_AT_DC'		WHEN '0XC000015B' THEN '15B_LOGON_TYPE_NOT_GRANTED'
				WHEN '0xC0000193' THEN '193_ACCOUNT_EXPIRED'			WHEN '0xC0000234' THEN '234_ACCOUNT_LOCKED_OUT'
				WHEN '0x0' THEN 'OK' END AS Status, 
			TO_UPPERCASE (extract_prefix(extract_suffix(TEXT, 0, 'logon of '), 0, '\\')) as Domain, 
			COUNT(*) AS Total 
		INTO $OutFile2
		FROM $InFiles
		WHERE 
			INDEX_OF(TO_UPPERCASE (TEXT),'SAMLOGON') >0 AND INDEX_OF(TO_UPPERCASE(TEXT),'RETURNS') >0 AND NOT INDEX_OF(TO_UPPERCASE(TEXT),'KERBEROS') >0 
		GROUP BY 
			Domain,Status ORDER BY Total DESC
"@
	Write-Progress -Activity "Generating $OutTitle2 report" -PercentComplete (60)
	$LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)| Out-Null
#--SamLogon-User_
	$OutTitle3 = 'SamLogon-User'
	$OutFile3 = $ScriptPath+'\'+$TimeStamp+'_'+$OutTitle3+'.csv'
	$Query = @"
		SELECT 
			CASE EXTRACT_SUFFIX(TEXT,0,'Returns ')
				WHEN '0XC000005E' THEN '5E_NO_LOGON_SERVERS' 			WHEN '0xC0000064' THEN '64_NO_SUCH_USER'
				WHEN '0xC000006A' THEN '6A_STATUS_WRONG_PASSWORD'		WHEN '0XC000006D' THEN '6D_LOGON_FAILURE'
				WHEN '0XC000006E' THEN '6E_ACCOUNT_RESTRICTION'			WHEN '0xC000006F' THEN '6F_INVALID_LOGON_HOURS'
				WHEN '0xC0000070' THEN '70_INVALID_WORKSTATION'			WHEN '0xC0000071' THEN '71_PASSWORD_EXPIRED'
				WHEN '0xC0000072' THEN '72_ACCOUNT_DISABLED'			WHEN '0XC00000DC' THEN 'DC_INVALID_SERVER_STATE'
				WHEN '0XC0000133' THEN '133_TIME_DIFFERENCE_AT_DC'		WHEN '0XC000015B' THEN '15B_LOGON_TYPE_NOT_GRANTED'
				WHEN '0xC0000193' THEN '193_ACCOUNT_EXPIRED'			WHEN '0xC0000234' THEN '234_ACCOUNT_LOCKED_OUT'
				WHEN '0x0' THEN 'OK' END AS Status, 
			TO_UPPERCASE (extract_prefix(extract_suffix(TEXT, 0, 'logon of '), 0, 'from ')) as User, 
			COUNT(*) AS Total
		INTO $OutFile3
		FROM $InFiles
		WHERE 
			INDEX_OF(TO_UPPERCASE (TEXT),'SAMLOGON') >0 AND INDEX_OF(TO_UPPERCASE(TEXT),'RETURNS') >0 AND NOT INDEX_OF(TO_UPPERCASE(TEXT),'KERBEROS') >0 
		GROUP BY 
			Status, User ORDER BY Total DESC
"@
	Write-Progress -Activity "Generating $OutTitle3 report" -PercentComplete (90)
	$LPQuery.ExecuteBatch($Query,$InputFormat,$OutputFormat)| Out-Null
#--
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($LPQuery) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($InputFormat) | Out-Null
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($OutputFormat) | Out-Null
#---------Find logs's time range Info----------
	$OldestEvent = [datetime]('12/31/2699')
	$NewestEvent = [datetime]('12/31/1999')
	$LogFiles = Get-ChildItem -Path $ScriptPath -Filter '*.log'
	foreach ($LogFile in $LogFiles) {
		$FirstLine = (Get-Content $LogFile -Head 1) -split ' '
		$LastLine  = (Get-Content $LogFile -Tail 1) -split ' '
			$FirstLineTime = [datetime]::ParseExact($FirstLine[0]+' '+$FirstLine[1],"MM/dd HH:mm:ss",$Null)
			$LastLineTime = [datetime]::ParseExact($LastLine[0]+' '+$LastLine[1],"MM/dd HH:mm:ss",$Null)
		If ($OldestEvent -ge $FirstLineTime) {$OldestEvent = $FirstLineTime }
		If ($NewestEvent -le $LastLineTime) {$NewestEvent = $LastLineTime }
	}
	$LogTimeRange = ($NewestEvent-$OldestEvent)
	$LogRangeText = 'Items: '+$OldestEvent+' ~ '+$NewestEvent+'; TimeRange = '+$LogTimeRange.Days+' Days, '+$LogTimeRange.Hours+' Hours, '+$LogTimeRange.Minutes+' Minutes, '+$LogTimeRange.Seconds+' Seconds'  
#---------Excel--------------------------------
	If (Test-Path $OutFile1) { # Check if LogParser generated CSV.
		$Excel = New-Object -ComObject excel.application  # https://docs.microsoft.com/en-us/office/vba/api/overview/excel/object-model
		$Excel.visible = $true
		#--SamLogon-Machine
			$Excel.Workbooks.OpenText("$OutFile1")
			CommonWorkBookTasks([int]1,[int]4,$OutTitle1,$LogRangeText) # 1-WorkBook 2-NumberColumnFormat 3-SheetTitle 4-RangeText
		#--SamLogon-Domain
			$Excel.Workbooks.Open($OutFile2) | Out-Null	
			CommonWorkBookTasks([int]2,[int]3,$OutTitle2,$LogRangeText)
		#--SamLogon-User
			$Excel.Workbooks.Open($OutFile3) | Out-Null	
			CommonWorkBookTasks([int]3,[int]3,$OutTitle3,$LogRangeText)

			[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
			# Stop-process -Name Excel 
		} else {
			Write-Host 'No LogParser result found. Please verify log type is Netlogon.log.' -ForegroundColor Red
		}
