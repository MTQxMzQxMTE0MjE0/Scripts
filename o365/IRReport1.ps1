<#
    date            2018/11/25
    purpose         o365 Incident Response Report 1of2
    requires        AzureAD, ExchangeOnline and MSOnline modules
#>

# Get Teanant domain name to use as file name
$TenantName = (Get-MsolDomain | Where-Object {$_.isDefault}).Name
$DomainName = $TenantName.Substring(0, $TenantName.IndexOf('.'))

# Variables
#$Now = (Get-Date -Format "yyyy-MM-dd-dddd")
$ReportPath = "$Env:USERPROFILE\Desktop\o365Reports\Daily\$DomainName"
#$ReportFile = "SpoofReport_" + $DomainName + "_" + $Now

# Report directory
If (!(Test-Path -Path $ReportPath)) {
    New-Item -Path $ReportPath -ItemType Directory -Force
}

# Spoof Report
Get-SpoofMailReport |
Select-Object Date,Action,CompAuthResult,CompAuthReason,SpfAuthStatus,DkimAuthStatus,DmarcAuthStatus,MessageCount,SpoofedSender,TrueSender,SendingInfrastructure,SenderIp |
Sort-Object -Property Date |
Export-Csv -Path $ReportPath\Spoof.csv -NoTypeInformation

# ATP Report
Get-MailTrafficATPReport |
Select-Object -Property Date,MessageCount,Direction,EventType,VerdictSource |
Sort-Object Date |
Export-Csv -Path $ReportPath\ATP.csv -NoTypeInformation

# Traffic summary
Get-ATPTotalTrafficReport -StartDate (Get-Date).AddDays(-30) -EndDate (Get-Date).AddDays(-1) |
Select-Object Date, EventType, MessageCount |
Export-Csv -Path $ReportPath\Summary.csv -NoTypeInformation

# Traffic report
Get-MailTrafficReport -Direction Inbound -EventType GoodMail -AggregateBy Hour |
Select-Object Date,EventType,Direction,MessageCount |
Export-Csv -Path $ReportPath\Aggregated-Inbound.csv

# Aggregated Outbound messages per hour
Get-MailTrafficReport -Direction Outbound -EventType GoodMail -AggregateBy Hour |
Select-Object Date,EventType,Direction,MessageCount |
Export-Csv -Path $ReportPath\Aggregated-Inbound.csv

# Tenant Report
$Reports = Get-ATPTotalTrafficReport -ErrorAction SilentlyContinue
$TenantReports = [PSCustomObject]@{
    TotalSafeLinkCount = ($Reports | where-object { $_.EventType -eq 'TotalSafeLinkCount' }).Messagecount
    TotalSpamCount     = ($Reports | where-object { $_.EventType -eq 'TotalSpamCount' }).Messagecount
    TotalBulkCount     = ($Reports | where-object { $_.EventType -eq 'TotalBulkCount' }).Messagecount
    TotalPhishCount    = ($Reports | where-object { $_.EventType -eq 'TotalPhishCount' }).Messagecount
    TotalMalwareCount  = ($Reports | where-object { $_.EventType -eq 'TotalMalwareCount' }).Messagecount
    DateOfReports      = "$($Reports.StartDate | Select-Object -Last 1) - $($Reports.EndDate | Select-Object -Last 1)"
}

# Create report file
$TenantReports |
Export-Csv -Path $ReportPath\Tenant.csv -NoTypeInformation
