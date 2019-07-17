function InitializeSCCM 
{ 
$ProcessMessage="`n Please wait.Initializing SCCM ........." 
 
# Site configuration
do
{ 
	write-host "`n Enter Site Code : " -foregroundcolor $inputcolor -nonewline 
	$SiteCode = read-host 
	$siteResult=($siteCode -match '\b^[a-zA-Z0-9]{3}\b')
	if(!$siteResult)
	{
	write-host " Site code can have only [3] alphanumeric characters. Please re-enter site code" -foregroundcolor RED
	}
}while(!$siteResult)

do
{
	write-host "`n Enter SMS Provider Server Name : " -foregroundcolor $inputcolor -nonewline 
	$ProviderMachineName = read-host 
	$nameResult=($ProviderMachineName -match '\b^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$\b')
	if(!$nameResult)
	{
	write-host " Entered SMS provider name is not valid as per naming conventions. Please re-enter provider name" -foregroundcolor RED
	}
	$connTest=Test-Connection -ComputerName $ProviderMachineName -Count 1 -ErrorAction SilentlyContinue
	if($connTest -eq $null)
	{
		write-host " Entered SMS provider is invalid or not responding. Please re-enter provider name" -foregroundcolor RED
	}
}while((!$nameResult) -or ($connTest -eq $null))



iex $ProcessColor 
sleep 2 
# Customizations 
$initParams = @{} 
 
# Import the ConfigurationManager.psd1 module  
if((Get-Module ConfigurationManager) -eq $null) { 
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams  
} 
 
# Connect to the site's drive if it is not already present 
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) { 
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams 
} 
 
# Set the current location to be the site code. 
Set-Location "$($SiteCode):\" @initParams 
return $true,$ProviderMachineName,$sitecode
} 

 
function deinitializeSCCM 
{ 
	$ProcessMessage="`n Please wait.De-Initializing SCCM ......" 
	iex $ProcessColor 
	sleep 2 
	$location="$env:SystemDrive"
	set-location $location 
} 
function updateHTML
{

param ($strPath)
IF(Test-Path $strPath)
  { 
   Remove-Item $strPath
   
  }

 }
 
 #--CSS formatting
$test=@'
<style type="text/css">
 h5,h2, th { text-align: left; font-family: Segoe UI;font-size: 13px;}
 h1{ text-align: left; font-family: Segoe UI;font-size: 20px;color:magenta;}
table { margin: left; font-family: Segoe UI; box-shadow: 10px 10px 5px #888; border: thin ridge grey;}
th { background: #4CAF50; color: #fff; max-width: 400px; padding: 5px 10px; font-size: 12px;}
td { font-size: 11px; padding: 5px 20px; color: #000; }
tr { background: #b8d1f3; }
tr:nth-child(even) { background: #f2f2f2; }
tr:nth-child(odd) { background: #ddd; }
</style>
'@

#--Variable declaration
 clear
 $location=get-location 
 $InputColor="yellow" 
 $ProcessColor="write-host `$ProcessMessage -ForegroundColor gray -BackgroundColor darkgreen" 
 $ReportTitle="SCCM Client Collection Membership"
 $strPath = "$location\$ReportTitle.html" 
 $report=0
 updateHTML $strPath
 $connectionResult=InitializeSCCM 
if($connectionResult[0])
{
#collection membership Information
do{
	clear
	$report++
		
			updateHTML $strPath
			$result =@()
			write-host "`n Enter computer name : " -foregroundcolor $inputcolor -nonewline
			$id=read-host
			$ResID = (Get-CMDevice -Name $id).ResourceID
			
			if($ResID -eq $null)
				{
					$ResID=(Get-WmiObject -ComputerName "$($connectionResult[1])" -Namespace  "ROOT\SMS\site_$($connectionResult[2])" -Query "SELECT * FROM SMS_R_System WHERE Name Like '$($ID)'").resourceid
				}
			[string]$systemName=$id.tostring().toupper()
			#$strPath = "$location\$systemName $ReportTitle$report.html" 
			$Collections = (Get-WmiObject -Class sms_fullcollectionmembership -Namespace "root\sms\site_$($connectionResult[2])" -ComputerName "$($connectionResult[1])" -Filter "ResourceID = '$($ResID)'").CollectionID
			foreach ($Collection in $Collections)
			{
				$result+=Get-CMDeviceCollection -CollectionId $Collection | select Name, CollectionID,MemberCount,LastMemberChangeTime 
			}
		
		if($result -ne $null)
			{
				$result =$result | Sort-Object -property Name
				ConvertTo-Html -Head $test -Title $ReportTitle -Body "<h1>  $ReportTitle </h1>" >>  "$strPath"
				ConvertTo-Html -Head $test -Title $ReportTitle -Body "<h1>  $systemName Collection membership</h1>" >>  "$strPath"
				$result| ConvertTo-html  -Head $test -Body "" >> "$strPath"
				write-host "`n Opening $strpath report. `n" -foregroundcolor $inputcolor -nonewline 
				Invoke-Item $strPath
			}
		else
			{
			
				write-host "`n System name entered is invalid."
			}
			
		do
			{
				$quit=read-host "`n Do you want to exit?[Type Yes or No]"
				$validyesno=($quit -match '^(?:Yes|No)$')
				if(!$validyesno)
				{
					write-host "`n Acceptable input [yes | no] only. Case insensitive." -foregroundcolor RED
				}
			}while(!$validyesno)
}while($quit -eq "No")
#De-initializing SCCM 
deinitializeSCCM 
}
else
{
write-host "Management point server is not responding or some error occurred during connection.Exiting......... " -foregroundcolor RED
}