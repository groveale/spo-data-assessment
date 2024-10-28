#######################################
# Description: This script will build a count of files that are open to all users in the tenant.
#              It uses MSGraph report API to pull SPO Site, M365 Groups and MS Team data. 
#              Group + Teams + SPO data can give us a public vs private file count across SPO.
#              Also possible to provide active vs inactive breakdown.
#
#              Auth required: Application permission - Reports.Read.All, Sites.Read.All
#
# Usage:       .\GetOpenFilesFromGroups.ps1
#
# Notes:       This script requires the MS Graph PowerShell module.
#              https://docs.microsoft.com/en-us/graph/powershell/installation
#
#  


##############################################
# Dependencies
##############################################

## Load the required modules

# try {
#     Import-Module Microsoft.Graph.Reports
#     Import-Module Microsoft.Graph.Sites
# }
# catch {
#     Write-Error "Error importing modules required modules - $($Error[0].Exception.Message))"
#     Exit
# }



##############################################
# Variables
##############################################

$clientId = "5cfc2462-cfc2-4c4c-a599-83308bb98165"
$tenantId = "75e67881-b174-484b-9d30-c581c7ebc177"
$thumbprint = "6ADC063641A24BB0BD68786AB71F07315CED9076"

$tempDataDir = ".\Temp"

##############################################
# Functions
##############################################

function ConnectToMSGraph 
{  
    try{
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint
    }   
    catch{
        Write-Host "Error connecting to MS Graph" -ForegroundColor Red
    }
}

function GetSPOReportData
{
    try 
    {
        Get-MgReportSharePointSiteUsageDetail -Period D180 -OutFile "$tempDataDir\SiteUsageDetail.csv"
    }
    catch
    {
        Write-Host "Error getting SPO report data - $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}

function GetTeamsReportData
{
    try 
    {
        Get-MgReportTeamActivityDetail -Period D180 -OutFile "$tempDataDir\TeamsDetail.csv"
    }
    catch
    {
        Write-Host "Error getting Teams report data - $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}

function GetGroupsReportData
{
    try 
    {
        Get-MgReportOffice365GroupActivityDetail -Period D180 -OutFile "$tempDataDir\GroupDetail.csv"
    }
    catch
    {
        Write-Host "Error getting Groups report data - $($Error[0].Exception.Message)" -ForegroundColor Red
    }
}

function GetGroupIdFromSiteId($siteId)
{
    try {
        $defaultDrive = Get-MgSiteDefaultDrive -SiteId $siteId -Select owner -ErrorAction Stop
        return $defaultDrive.Owner.AdditionalProperties.group.id  
    }
    catch {
        Write-Host "Error getting group ID for site ID: $siteId - $($Error[0].Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function GetGroupVisibiliy($groupId, $teamsData)
{
    if ($null -eq $groupId)
    {
        return $null
    }
    return $teamsData | Where { $_.'Group Id' -eq $groupId } | Select -ExpandProperty "Group Type"
}

##############################################
# Main
##############################################

## Auth
ConnectToMSGraph

## Create temp dir
if (-not (Test-Path $tempDataDir))
{
    New-Item -Path $tempDataDir -ItemType Directory
}

## Get Report data
# GetSPOReportData
# GetTeamsReportData
GetGroupsReportData

## Read and transform data
$siteData = Import-Csv "C:\Users\alexgrover\source\repos\spo-data-assessment\tempspodetail.csv"
$teamsData = Import-Csv "teamsdetail.csv"
$groupsData = Import-Csv "$tempDataDir\GroupDetail.csv"

## We have x sites and z teams
Write-Host "Site data count: $($siteData.Count)"
Write-Host "Teams data count: $($teamsData.Count)"
Write-Host "Groups data count: $($groupsData.Count)" 

## iterate over the site data and get the group ID for each site
$groupConnectedSites = $siteData | Where { $_.'Root Web Template' -eq "Group" }

## We have y sites connected to a group
Write-Host "Group connected sites count: $($groupConnectedSites.Count)"

## Get the group ID for each site
$groupConnectedSites | % {
    $groupId = GetGroupIdFromSiteId($_.'Site Id')
    $_ | Add-Member -MemberType NoteProperty -Name GroupId -Value $groupId
    $_ | Add-Member -MemberType NoteProperty -Name Visibility -Value (GetGroupVisibiliy $groupId $groupsData)
}

## Count sites that have a group ID
$groupConnectedSitesWithGroupId = $groupConnectedSites | Where { $_.Visibility -ne $null }

## We have z sites with a group ID
Write-Host "Sites with a connected team: $($groupConnectedSitesWithGroupId.Count)"



