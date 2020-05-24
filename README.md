# Fix-PublicFolderNames.ps1

Script to prepare legacy public folder names for migration to modern public folders

## Description

This script renames legacy public foldern on Exchange Server 2010 to replace backslash "\" and forward slash "/" by the pipe "|" character.

This script trims all public folder names to remove any leading or trailing spaces.

## Requirements

- Windows Server 2008 R2 SP1, Windows Server 2012 or Windows Server 2012 R2  
- Exchange 2010 Management Shell

## Parameters

### PublicFolderServer

Exchange Server name hosting legacy public folders

## Examples

``` PowerShell
.\Fix-PublicFolderNames -PublicFolderServer MYPFSERVER
```

Rename and trim public folders found on Server MYPFSERVER

## Note

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE
RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

## Credits

Written by: Thomas Stensitzki

## Stay connected

- My Blog: [http://justcantgetenough.granikos.eu](http://justcantgetenough.granikos.eu)
- Twitter: [https://twitter.com/stensitzki](https://twitter.com/stensitzki)
- LinkedIn: [http://de.linkedin.com/in/thomasstensitzki](http://de.linkedin.com/in/thomasstensitzki)
- Github: [https://github.com/Apoc70](https://github.com/Apoc70)
- MVP Blog: [https://blogs.msmvps.com/thomastechtalk/](https://blogs.msmvps.com/thomastechtalk/)
- Tech Talk YouTube Channel (DE): [http://techtalk.granikos.eu](http://techtalk.granikos.eu)

For more Office 365, Cloud Security, and Exchange Server stuff checkout services provided by Granikos

- Blog: [http://blog.granikos.eu](http://blog.granikos.eu)
- Website: [https://www.granikos.eu/en/](https://www.granikos.eu/en/)
- Twitter: [https://twitter.com/granikos_de](https://twitter.com/granikos_de)