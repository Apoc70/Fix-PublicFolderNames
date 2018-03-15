<#
    .SYNOPSIS
    Script to prepare legacy public folder names for migration to modern public folders
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0, 2018-03-15

    Ideas, comments and suggestions to support@granikos.eu 
 
    .LINK  
    http://scripts.Granikos.eu
	
    .DESCRIPTION
	
    This script renames legacy public foldern on Exchange Server 2010 to replace backslash "\" and forward slash "/" by the pipe "|" character.

    This script trims all public folder names to remove any leading or trailing spaces.

    .NOTES 
    Requirements 
    - Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
    - Exchange 2010 Management Shell

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
    
    .PARAMETER PublicFolderServer
    Exchange Server name hosting legacy public folders
    
    .EXAMPLE
    Rename and trim public folders found on Server MYPFSERVER

    .\Fix-PublicFolderNames -PublicFolderServer MYPFSERVER


#>

[CmdletBinding()]
Param(
  [string] $PublicFolderServer = 'PFSERVER01'
)

Write-Host ('Fetching Public Folder Statistic on Server {0}' -f ($PublicFolderServer))

# Fetch Public Folders containing unsupported characters
$FilteredFolders = Get-PublicFolderStatistics -ResultSize Unlimited -server $PublicFolderServer | Where-Object {($_.Name -like "*\*")-or ($_.Name -like "*/*") } 

Write-Host ('Found {0} public folders' -f ($FilteredFolders | Measure-Object).Count)

foreach($Folder in $FilteredFolders) {
  $Name = $Folder.Name.Replace('\','|').Replace('/','|')
  Write-Host ('Fixing folder: [{0}] > [{1}]' -f $Folder.Name, ($Name))
  Set-PublicFolder -Server $PublicFolderServer -Identity $Folder.Identity -Name $Name
}

# Now trim all public folders
$AllPublicFoldersFolders = Get-PublicFolderStatistics -ResultSize Unlimited -server $PublicFolderServer 
$AllPublicFoldersFolders | ForEach-Object { Set-PublicFolder -Identity $_.Identity -Name $_.Name.Trim() }