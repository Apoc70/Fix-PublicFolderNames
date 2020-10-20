<#
    .SYNOPSIS
    Script to prepare on-premises modern public folder names for migration to modern public folders in Exchange Online 
   
    Thomas Stensitzki
	
    THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE 
    RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.
	
    Version 1.0, 20202-10-20

    Use the GitHub repository for ideas, comments, and suggestions.
 
    .LINK  
    http://scripts.Granikos.eu
	
    .DESCRIPTION	
    This script renames modern public folder names an replaces unsupported characters with the hyphen "-" character.

    This script trims public folder names to remove any leading or trailing spaces.

    .NOTES 
    Requirements 
    - Windows Server Windows Server 2012R2 or newer
    - Exchange 2013 Management Shell or newer
    - Organization Management RBAC Management Role (let's keep it simple)

    Revision History 
    -------------------------------------------------------------------------------- 
    1.0     Initial community release 
        
    .PARAMETER ExportFolderNames
    Switch to export renamed folders to text files
    
    .EXAMPLE
    Rename and trim public folders

    .\Fix-ModernPublicFolderNames 

    .EXAMPLE
    Rename and trim public folders, export list of renamed folders and folders with renaming errors as text file

    .\Fix-ModernPublicFolderNames -ExportFolderNames

#>

[CmdletBinding()]
Param(
  [switch] $ExportFolderNames
)

Write-Host 'Fetching Public Folders'

# some variables
$FolderScope = '\'
$ScriptDir = Split-Path -Path $script:MyInvocation.MyCommand.Path
$TimeStamp = $(Get-Date -Format 'yyyy-MM-dd HHmm')
$FileNameSuccess = 'RenamedFoldersSuccess'
$FileNameError = 'RenamedFoldersError'

# Fetch all public folders
$PublicFolders = Get-PublicFolder $FolderScope -Recurse -ResultSize Unlimited

# Fetch Public Folders containing unsupported characters
$FilteredFolders = $PublicFolders | Where-Object { ($_.Name -like "*\*") -or ($_.Name -like "*/*") -or ($_.Name -like "*;*")-or ($_.Name -like "*:*")-or ($_.Name -match ",")} 

$UpdatedFolders = @()
$FoldersWithError = @()
$PublicFolderCount = ($FilteredFolders | Measure-Object).Count

Write-Host ('Found {0} public folders with names containing unsupported characters' -f $PublicFolderCount)

$Count=0

foreach ($PublicFolder in $FilteredFolders) { 
  
  # write some nice progress bar
  Write-Progress -Activity "Replace characters" -Status "Replace characters: $([math]::Round($(($Count/$PublicFolderCount)*100))) %" -PercentComplete (($Count/$PublicFolderCount)*100) -SecondsRemaining $($PublicFolderCount-$Count)
  
  # build new folder name by replacing unknown characters and trimming the name
  $NewPublicFolder = ($PublicFolder.Name -replace ("\\|\/|\:|\;|\,","-")).Trim()
  
  try { 
    # try to set the new folder name
    Set-PublicFolder -Identity $PublicFolder.EntryId -name $NewPublicFolder -Confirm:$false  -EA 'stop' 
    $UpdatedFolders += $NewPublicFolder
   } 
   catch {
    # Ooops 
    Write-Host $Error[0].Exception.Message -F yellow 
    $FoldersWithError += $PublicFolder
  }
  
  $Count++
}

if($ExportFolderNames) {
  # Write successfully renamed folders to file
  $OutputFile = (Join-Path -Path $ScriptDir -ChildPath ('{0}-{1}.txt' -f $FileNameSuccess, $TimeStamp))
  $UpdatedFolders | Out-File -FilePath $OutputFile -Force -Confirm:$false
  
  # Write failed folders to file
  $OutputFile = (Join-Path -Path $ScriptDir -ChildPath ('{0}-{1}.txt' -f $FileNameError, $TimeStamp))
  $FoldersWithError | Out-File -FilePath $OutputFile -Force -Confirm:$false 
}

Write-Host ('Folders updated successfully: {0}' -f ($UpdatedFolders | Measure-Object).Count)
Write-Host ('Folders with update errors  : {0}' -f ($FoldersWithError | Measure-Object).Count)