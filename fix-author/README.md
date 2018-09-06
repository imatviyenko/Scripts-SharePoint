# fix-author.ps1
PowerShell script for fixing 'User cannot be found' SharePoint errors when accessing files saved in document libraries in cases when these errors are caused by renamed/deleted Windows logins of the file authors.

## Description
When a user saves a file in a SharePoint document library, his current login is saved as a string in the properties of the file. If his/her login name is changed or the user is removed from the list of the web users, certain operations on this file, such as getting the file's author or reading this file from external custom code using the CSOM library, may throw an exception 'User cannot be found'. This happens when the current login name of the file's author has changed for some reasons, or the user was completely removed from the SharePoint.
This script can be used to find all such problematic files and change the file author attribute to some valid login name in 'fix' mode, or just output the report to the screen in 'report' mode.
Usually you would like to run the script in 'report' mode on a certain folder in a document library first to get the list of the missing/incorrect authors logins and then use this list to prepare a hashtable with the mappings "old login"->"new login" required for the 'fix' mode.

## Usage
Use the built-on PowerShell Get-Help facility to get information on usage scenarios as well as some examples.
**PS C:\>Get-Help .\fix-authors.ps1**