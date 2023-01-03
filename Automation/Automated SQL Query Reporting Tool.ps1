### FILL THESE IN
# SQL Command Variables
$ServerInstance = "SW18V560"
$Database = "GSAUProd"
$Query = ""

# Variables for file saving
$OutputFolder = "$PSScriptRoot\Output\"
$ArchiveFolder = "$PSScriptRoot\Archive\"

# Set filename to scriptname
$scriptName = $MyInvocation.MyCommand.Name
$filename = $scriptName.Substring(0,$scriptName.Length-4)+" "+ (Get-Date -Format "yyyy-MM")
# Email for report failure notification
$recipients = 'dheydon@groupeseb.com'
# Log File Full Path Name
$LogFile = "$PSScriptRoot\"+$scriptName.Substring(0,$scriptName.Length-4)+".Log"

# Log Method
function Log {
    param( [Parameter(ValueFromPipeline)][String]$msg )
    if (!(Test-Path $LogFile)){
           New-Item -path "$PSScriptRoot\" -name ($scriptName.Substring(0,$scriptName.Length-4)+".Log") -type "file"
           Add-Content $LogFile $msg
        }
    else{
           Add-Content $LogFile $msg
        }
}



# Standardised Query Method
function runQuery {
param(
    [Parameter()] 
    $fileName,
    [Parameter()] 
    $server,
    [Parameter()] 
    $Database,
    [Parameter()]
    $query
)

$csv = $PSScriptRoot +'\'+ $filename +  ".csv"
$xlsx = $PSScriptRoot +'\'+ $filename +  ".xlsx"

# Log remove xlsx
((Get-Date -Format "yyyy-MM-dd hh:mm:ss")+ ": Removing $xlsx")| Log
Remove-Item $xlsx

Start-Sleep -Seconds 10

# Log Start
((Get-Date -Format "yyyy-MM-dd hh:mm:ss")+ ": Begin run query")| Log
$result = Invoke-Sqlcmd -QueryTimeout 0 -Query $query -Database $Database -ServerInstance $server
$result |  Export-Csv -Path $csv -NoTypeInformation
## Export to xlsx using excel (Preserves Date)
Start-Process -FilePath 'C:\Program Files (x86)\Microsoft Office\root\Office16\excelcnv.exe' -ArgumentList "-nme -oice ""$csv"" ""$xlsx"""
Start-Sleep -Seconds 10

Remove-item $csv

# Test for query success
    IF (Test-Path $xlsx) {
        If ((Get-Item $xlsx).length -gt 10kb) {
            # Save the file in the output folder
            Copy-item $xlsx -Destination $OutputFolder
            ((Get-Date -Format "yyyy-MM-dd hh:mm:ss")+": Copy Excel file to " + $OutputFolder)| Log
            # Archive the file
            ((Get-Date -Format "yyyy-MM-dd hh:mm:ss")+": Archiving excel file")| Log
            Move-item $xlsx -Destination $ArchiveFolder
        }
        Else {
            # Query has failed send email to notify
            ((Get-Date -Format "yyyy-MM-dd hh:mm:ss")+ ": Query failed, report size LT 10KB") | Log
            $EmailString = ("$filename failed to run, the script is saved on " + ($PSScriptRoot).Replace("C:","\\SW49T369\C$"))
            $From = 'gsauhelpdesk@groupeseb.com'
            $Date = Get-Date -Format â€œdd.MM.yyyyâ          
            Send-MailMessage -From $From -To $recipients -Subject ((Get-Date -Format "yyyy-MM-dd hh:mm:ss")+ ": $filename failed") -Body $EmailString -SmtpServer 'smtp.seb.com' 
        }
    }
}

#  Run the query
runQuery -fileName $filename -server $ServerInstance -Database $Database -query $Query





