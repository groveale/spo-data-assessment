#######################################
# Description: This script will build a count of files that are open to all users in the tenant.
#              It uses MSGraph report API to pull SPO Site, M365 Groups and MS Team data. 
#              Group + Teams + SPO data can give us a public vs private file count across SPO.
#              Also possible to provide active vs inactive breakdown.
#
#              Auth required: Application permission - Reports.Read.All, Drive.Read.All
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

try {
    Import-Module Microsoft.Graph.Beta.Reports
    Import-Module Microsoft.Graph.Files
}
catch {
    Write-Error "Error importing modules required modules - $($Error[0].Exception.Message))"
    Exit
}



##############################################
# Variables
##############################################

$clientId = "5cfc2462-cfc2-4c4c-a599-83308bb98165"
$tenantId = "75e67881-b174-484b-9d30-c581c7ebc177"
$thumbprint = "6ADC063641A24BB0BD68786AB71F07315CED9076"

$today = Get-Date

$tempDataDir = ".\Temp"

# Define the path to the HTML file
$htmlFilePath = "dashboard.html"

# Sites and Teams will be considered inactive if they have not been active in the last 180 days
$inactiveDays = 365

##############################################
# Functions
##############################################

function ConnectToMSGraph 
{  
    try{
        Connect-MgGraph -ClientId $clientId -TenantId $tenantId -CertificateThumbprint $thumbprint -NoWelcome
    }   
    catch{
        Write-Host "Error connecting to MS Graph" -ForegroundColor Red
    }
}

function GetSPOReportData
{
    try 
    {      
        ## Check if file exists at path and if so exit
        if (Test-Path "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-SiteUsageDetail.csv")
        {
            Write-Host "SPO report data already exists for today. Skipping data pull." -ForegroundColor Yellow
            return
        }  
        Get-MgBetaReportSharePointSiteUsageDetail -Period D180 -OutFile "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-SiteUsageDetail.csv"
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
        ## Check if file exists at path and if so exit
        if (Test-Path "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-TeamsDetail.csv")
        {
            Write-Host "Teams report data already exists for today. Skipping data pull." -ForegroundColor Yellow
            return
        }

        Get-MgReportTeamActivityDetail -Period D180 -OutFile "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-TeamsDetail.csv"
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
        ## Check if file exists at path and if so exit 
        if (Test-Path "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-GroupDetail.csv")
        {
            Write-Host "Group report data already exists for today. Skipping data pull." -ForegroundColor Yellow
            return
        }

        Get-MgReportOffice365GroupActivityDetail -Period D180 -OutFile "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-GroupDetail.csv"
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

function CalcualteTotals($data, $siteData)
{
    # Initialize counters
    $totals = @{
        sites = 0
        publicSites = 0
        teams = 0
        publicTeams = 0
        files = 0
        activeFiles = 0
        publicFiles = 0
        activePublicFiles = 0
        inactivePublicSites = 0
        inactivePublicTeams = 0
        inactivePublicFiles = 0
        orgWideLinks = $siteData.'Company Link Count' | Measure-Object -Sum | Select -ExpandProperty "Sum"
        anonLinks = $siteData.'Anonymous Link Count' | Measure-Object -Sum | Select -ExpandProperty "Sum"
        specificLinks = $siteData.'Secure Link For Member Count' | Measure-Object -Sum | Select -ExpandProperty "Sum"
        inactiveDays = $inactiveDays
    }

    # Iterate through each row and update the counters
    foreach ($row in $data) {
        if ($row."Teams Connected" -eq "True") {
            $totals.teams++
            if ($row.Visibility -eq "Public") {
                $totals.publicTeams++
                if (([datetime]$row.'Last Activity Date' -lt (Get-Date).AddDays(-$inactiveDays))) {
                    $totals.inactivePublicTeams++
                    $totals.inactivePublicFiles += [int]$row."File Count"
                }
            }
            
        }
        else 
        {
            $totals.sites++
            if ($row.Visibility -eq "Public") {
                $totals.publicSites++
                if ([datetime]$row.'Last Activity Date' -lt (Get-Date).AddDays(-$inactiveDays)) {
                    $totals.inactivePublicSites++
                    $totals.inactivePublicFiles += [int]$row."File Count"
                }
            }
        }
        
        $totals.files += [int]$row."File Count"
        $totals.activeFiles += [int]$row."Active File Count"
        if ($row.Visibility -eq "Public") {
            $totals.publicFiles += [int]$row."File Count"
            $totals.activePublicFiles += [int]$row."Active File Count"
        }
    }

    return $totals
}

function PopulateHTML($totals)
{
    # Read the content of the HTML file
    $htmlContent = Get-Content -Path $htmlFilePath -Raw

    # Define the new data for dashboardData
    $newDashboardData = @"
    <script>
        window.dashboardData = {
            files: $($totals.files),
            publicSites: $($totals.publicSites),
            publicFiles: $($totals.publicFiles),
            publicTeams: $($totals.publicTeams),
            sites: $($totals.sites),
            activePublicFiles: $($totals.activePublicFiles),
            teams: $($totals.teams),
            activeFiles: $($totals.activeFiles),
            inactivePublicSites: $($totals.inactivePublicSites),
            inactivePublicTeams: $($totals.inactivePublicTeams),
            inactivePublicFiles: $($totals.inactivePublicFiles),
            orgWideLinks: $($totals.orgWideLinks),
            anonLinks: $($totals.anonLinks),
            specificLinks: $($totals.specificLinks),
            daysInactive: $($totals.inactiveDays)
        };
    </script>
"@

    # Use a regular expression to replace the existing dashboardData script block
    $updatedHtmlContent = [regex]::Replace($htmlContent, '<script>\s*window\.dashboardData\s*=\s*\{[^}]+\};\s*</script>', $newDashboardData)

    # Write the updated content back to the HTML file
    Set-Content -Path $htmlFilePath -Value $updatedHtmlContent

    Write-Output "Dashboard data updated successfully."
}

function PopulateDataFrame($siteData, $groupsData, $teamsData)
{

    # Initialize an empty hashtable
    $siteDataHashtable = @{}

    # Iterate over each item in $siteData
    # foreach ($site in $siteData) {
    #     # Assuming the site ID is stored in a property called 'SiteID'
    #     $siteID = $site.'Site Id'

    #     # Add a new property to the object called 'Group Id' and set it to the value returned by GetGroupIdFromSiteId
    #     $site | Add-Member -MemberType NoteProperty -Name "Group Id" -Value (GetGroupIdFromSiteId $siteID)

    #     # Add the item to the hashtable using the site ID as the key
    #     $siteDataHashtable[$siteID] = $site
    # }

    ## filter out group connected sites
    $nonGroupSites = $siteData | Where { $_.'Root Web Template' -ne "Group" }

    $groupConnectedSites = $siteData | Where { $_.'Root Web Template' -eq "Group" } 


    ## We apepar to have some group connected team sites that are have a root web template of "Team Site"
    ## We will attempt to filter these out using the owner property of the site. If the Owner ends in "Owners" then it highly likey a group connected site.
    $nonGroupSitesReal = $nonGroupSites | Where { !($_.'Root Web Template' -eq "Team Site" -and $_.'Owner Display Name'.EndsWith("Owners")) }

    Write-Host "SPO Sites (non-group) count: $($nonGroupSitesReal.Count)"

    ## non group sites + groups == total sites
    # $nonGroupSitesReal.Count + $groupsData.Count
    # $siteData.Count

    ## teamGroupIds
    $teamsGroupId = $teamsData.'Team Id'

    ## Create a new data frame with the SiteId, GroupId, DataSource (SPO, Group), Visibility (Public, Private), Last Activity Date, File Count, Active File Count, Teams Connected
    $dataFrame = @()

    # Iterate through the first list (nonGroupSitesReal)
    foreach ($site in $nonGroupSitesReal) {
        $dataFrame += [PSCustomObject]@{
            SiteId            = $site.'Site Id'
            GroupId           = [string]::Empty
            DataSource        = "SPO"
            Visibility        = "Private" # Assumed Private but can have the EEEU claim - MGDC required
            'Last Activity Date' = if ([string]::Empty -eq $site.'Last Activity Date') { "1970-01-01" } else { $site.'Last Activity Date' }
            'File Count'      = $site.'File Count'
            'Active File Count' = $site.'Active File Count'
            'Teams Connected' = $false
            'Owner Principal Name' = $site.'Owner Principal Name'
            'Org Wide Link Count' = $site.'Company Link Count'
            'Specific People Link Count' = $site.'Secure Link For Member Count'
            'Anonymous Link Count' = $site.'Anonymous Link Count'
            'Group Members Count' = 0
        }
    }

    # Iterate through the second list (groupsData)
    foreach ($group in $groupsData) {

        #$siteId = $groupSites | Where { $_.'Group Id' -eq $group.'Group Id' } | Select -ExpandProperty "Site Id"

        $dataFrame += [PSCustomObject]@{
            SiteId            = [string]::Empty
            GroupId           = $group.'Group Id'
            DataSource        = "Group"
            Visibility        = $group.'Group Type'
            'Last Activity Date' = if ([string]::Empty -eq $group.'Last Activity Date') { "1970-01-01" } else { $group.'Last Activity Date' }
            'File Count'      = if ([string]::Empty -eq $group.'SharePoint Total File Count') { 0 } else { $group.'SharePoint Total File Count' }
            'Active File Count' = if ([string]::Empty -eq $group.'SharePoint Active File Count') { 0 } else { $group.'SharePoint Active File Count' }
            'Teams Connected' = $teamsGroupId.Contains($group.'Group Id')
            'Owner Principal Name' = $group.'Owner Principal Name'
            'Org Wide Link Count' = [string]::Empty
            'Specific People Link Count' = [string]::Empty
            'Anonymous Link Count' = [string]::Empty
            'Group Members Count' = $group.'Member Count'
        }
    }
    return $dataFrame
}

##############################################
# Main
##############################################

## Auth
#ConnectToMSGraph

## Create temp dir
if (-not (Test-Path $tempDataDir))
{
    New-Item -Path $tempDataDir -ItemType Directory
}

## Get Report data
# GetSPOReportData
# GetTeamsReportData
# GetGroupsReportData

## Read and transform data
# $siteData = Import-Csv "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-SiteUsageDetail.csv"
# $teamsData = Import-Csv "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-TeamsDetail.csv"
# $groupsData = Import-Csv "$tempDataDir\$($today.ToString("yyyy-MM-dd"))-GroupDetail.csv"

## remove any deleted rows
$siteData = $siteData | Where { $_.'Is Deleted' -eq $false }
$teamsData = $teamsData | Where { $_.'Is Deleted' -eq $false }
$groupsData = $groupsData | Where { $_.'Is Deleted' -eq $false }

## We have x Teams and z groups
Write-Host "Teams data count: $($teamsData.Count)"
Write-Host "Groups data count: $($groupsData.Count)" 
Write-Host "SPO Sites count: $($siteData.Count)"

$dataFrame = PopulateDataFrame $siteData $groupsData $teamsData

## Write to CSV
$dataFrame | Export-Csv -Path ".\Output\OpenFilesData.csv" -NoTypeInformation

## Calculate totals
$totals = CalcualteTotals -data $dataFrame -siteData $siteData

## Populate HTML
PopulateHTML($totals)

$totals

## Launch HTML
Start-Process -FilePath $htmlFilePath

##############################################
# End


