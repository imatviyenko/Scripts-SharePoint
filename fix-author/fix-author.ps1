<#
.SYNOPSIS
Script for fixing 'User cannot be found' SharePoint errors when accessing files saved in document libraries in cases when these errors are caused by renamed/deleted Windows logins of the file authors


.DESCRIPTION
When a user saves a file in a SharePoint document library, his current login is saved as a string in the properties of the file. If his/her login name is changed or the user is removed from the list of the web users, certain operations on this file, such as geting the file's author or reading this file from external custom code using the CSOM library, may throw an exception 'User cannot be found'. This happens when the current login name of the file's author has changed for some reasons, or the user was completelly removed from the Sharepoint.
This script can be used to find all such problematic files and change the file author attribute to some valid login name in 'fix' mode, or just output the report to the screen in 'report' mode.
Usually you would like to run the script in 'report' mode on a certain folder in a document library first to get the list of the missig/incorrect authors logins and then use this list to prepare a hashtable with the mappings "old login"->"new login" required for the 'fix' mode.

.PARAMETER WebUrl
Full URL for of the SPWeb where the document library is hosted

.PARAMETER RootFolderPath
The path to the document library root folder from which the processing is started

.PARAMETER ProcessSubfolders
Optional switch parameter specifying whether to process files (default) or subfolders in the root folder

.PARAMETER Mode
Processing mode, can be either 'report' for reporting only or 'fix' for actual replacement of old invalid logins with the new ones according to the mappings provided in the 'LoginMappings' hashtable parameter

.PARAMETER StdOut
Script output mode, specifies whether to use colored output to console (default) or use the standard output stream if you add this switch parameter, which can be further redirected to a text file


.PARAMETER LoginMappings
Hashtable with the mappings of old logins to new logins for file authors. Required if 'fix' value was specified for the 'Mode' parameter. The logins must be in the claims format (i.e., i:0#.w|DOMAIN\samAccountName)


.EXAMPLE
.\fix-author.ps1 -WebURL "https://sp.company.com/sites/it" -RootFolderPath "Inventory" -Mode report

This command will check all files in the "Inventory" document libary under the "https://sp.company.com/sites/it" SharePoint web site for missing/incorrect logins of file authors and will output the report including the number of 'bad' files with such an author and the list of all found problematic file authors' logins


.EXAMPLE
.\fix-author.ps1 -WebURL "https://sp.company.com/sites/it" -RootFolderPath "Inventory" -Mode report -StdOut | Out-File .\scriptlog.txt

The same as the previous example, but all script output is sent to the standard output stream and redirected to a text file 'scriptlog.txt' in the current folder


.EXAMPLE
.\fix-author.ps1 -WebURL "https://sp.company.com/sites/sales" -RootFolderPath "InvoiceAttachments/2018" -ProcessSubfolder -Mode report

This command will check all subfolders (note the -ProcessSubfolder switch parameter) in the folder "2018" inside the "InvoiceAttachments" document libary under the "https://sp.company.com/sites/sales" SharePoint web site for missing/incorrect logins of file authors. The of the '-ProcessSubfolders' option can be useful if you have a document library which is used for storing multiple attachments per one 'master' SharePoint form/list item, where for example all attachments for the 'master' item with ID 1234 for ther current year are stored in the 'InvoiceAttachments/2018/1234' subfolder.


.EXAMPLE
$mappings = @{"i:0#.w|COMPANY\old.login1" = "i:0#.w|COMPANY\new.login1"; "i:0#.w|COMPANY\old.login2" = "i:0#.w|COMPANY\new.login2"};
.\fix-author.ps1 -WebURL "https://sp.company.com/sites/it" -RootFolderPath "Inventory" -Mode fix -LoginMappings $mappings

This command will go through all files in the "Inventory" document libary under the "https://sp.company.com/sites/it" SharePoint web site and replace "author" and "modifiedby" fields on all ocassions where the current value of this field is equal to one of the 'old' logins (key column) in the $mappings hashtable, using the hashtable lookup value for the replacement.
Please note the use of the claims format for Windows logins (i.e., i:0#.w|DOMAIN\samAccountName).


.NOTES
Author: Ivan Matviyenko
Date:   September 5, 2018

.LINK
https://imatviyenko.github.io

#>


[CmdletBinding()]

Param(
  [Parameter(
    Position=1,
    Mandatory=$true,
    HelpMessage="Full URL for of the SPWeb where the document library is hosted"
  )]
  [string] $WebURL,

  [Parameter(
    Position=2,
    Mandatory=$true,
    HelpMessage="The path to the document library root folder from which the processing is started"
  )]
  [string] $RootFolderPath,

  [Parameter(
    Position=3,
    Mandatory=$false,
    HelpMessage="Switch parameter specifying whether to process files or subfolders in the root folder"
  )]
  [switch] $ProcessSubfolders,

  [Parameter(
    Position=4,
    Mandatory=$true,
    HelpMessage="Processing mode, can be either 'report' for reporting only or 'fix' for actual replacement of old invalid logins with the new ones according to the mappings provided in the 'LoginMappings' hashtable parameter"
  )]
  [ValidateSet("report","fix")]
  [string] $Mode = "report",

  [Parameter(
    Position=5,
    Mandatory=$false,
    HelpMessage="Switch parameter specifying whether to use colored output to console (default) or use the standard output stream"
  )]
  [switch] $StdOut
)


DynamicParam {
    if ($Mode -eq "fix") {
        $loginMappingsParamAttribute = New-Object System.Management.Automation.ParameterAttribute;
        $loginMappingsParamAttribute.HelpMessage = "Hashtable with the mappings of old logins to new logins for file authors";
        $loginMappingsParamAttribute.Mandatory = $true;
        $loginMappingsParamAttribute.Position = 6;
        $attributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute];
        $attributeCollection.Add($loginMappingsParamAttribute);
        $loginMappingsParam = New-Object System.Management.Automation.RuntimeDefinedParameter('LoginMappings', [hashtable], $attributeCollection);
        $paramDictionary = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary;
        $paramDictionary.Add('LoginMappings', $loginMappingsParam);
        return $paramDictionary;
    }
}

Process {
    $ErrorActionPreference = "Stop";
    Add-PSSnapin Microsoft.SharePoint.PowerShell;


    Write-Host "`n*************************************************************";
    Write-Host "Starting script in $mode mode`n`n";

    if ($Mode -eq "fix") {
        Write-Verbose "PSBoundParameters.LoginMappings: $($PSBoundParameters.LoginMappings)";
        $Script:loginMappings = $PSBoundParameters.LoginMappings;
    };
    

    $spWeb = Get-SPWeb $WebUrl;
    $spRootFolder = $spWeb.GetFolder($RootFolderPath);

    if ($Mode -eq "report") {
        $Script:goodFolderCount = 0;
        $Script:badFolderCount = 0;
        $Script:invalidAuthorLogins = @();
    } elseif ($Mode -eq "fix") {
        $Script:fixedFolderCount = 0;
        $Script:unfixedFolderCount = 0;
    }
    

    $counter = 0;

    if ($ProcessSubfolders -eq $false) {
        if ($Mode -eq "report") {
            ReportDocLibFolder $spRootFolder;
        } elseif ($Mode -eq "fix") {
            FixDocLibFolder $spRootFolder;
        };
        $counter++;
    } else {
        foreach ($spFolder in $spRootFolder.SubFolders) {
            $counter++;
            if ($Mode -eq "report") {
                ReportDocLibFolder $spFolder;
            } elseif ($Mode -eq "fix") {
                FixDocLibFolder $spFolder;
            };
        };
    }

    Write-Host "`n`nCount of processed folders: $counter";
    if ($Mode -eq "report") {
        Write-Host "goodFolderCount: $Script:goodFolderCount";
        Write-Host "badFolderCount: $Script:badFolderCount";
        $invalidLoginsAsString = $Script:invalidAuthorLogins -join ";";
        Write-Host "`ninvalidAuthorLogins: $invalidLoginsAsString";
    } else {
        Write-Host "fixedFolderCount: $Script:fixedFolderCount";
        Write-Host "unfixedFolderCount: $Script:unfixedFolderCount";
    };

    Write-Host "`nScript completed execution";
    Write-Host "*************************************************************";
}

Begin {
        
        $colorMappings = @{"INFO" = "green"; "WARNING" = "yellow"; "ERROR" = "red"};
        function WriteOutput($level, $status, $message){
            if ($StdOut) {
                Write-Output "$level`t$status`t$message";
            } else {
                $color = $colorMappings[$level];
                Write-Host "$level`t$status`t$message" -ForegroundColor $color;
            }
        }
        
        
        function FixAuthor($spFile, $badLogin) {
            Write-Verbose "FixAuthor invoked";
            Write-Verbose "spFile.Name: $($spFile.Name)";
            Write-Verbose "Script:loginMappings: $($Script:loginMappings)";
            $spFileItem = $spFile.Item;
            $goodLogin = $Script:loginMappings[$badLogin];
            if ($goodLogin -eq $null) {
                return $null; # no mapping found
            };
    
    
            $createdDate = $spFileItem["Created"];
            $modifiedDate = $spFileItem["Modified"];
            Write-Verbose "createdDate: $createdDate `t modifiedDate: $modifiedDate";

            [bool] $dirty = $false;

            if ($spFileItem.Properties["vti_author"] -eq $badLogin) {
                Write-Verbose "badLogin: $badLogin `t goodLogin: $goodLogin";
                $spFileItem.Properties["vti_author"] = $goodLogin;
                $dirty = $true;
            };

            if ($spFileItem.Properties["vti_modifiedby"] -eq $badLogin) {
                $spFileItem.Properties["vti_modifiedby"] = $goodLogin;
                $dirty = $true;
            };

            if ($dirty) {
                $spFileItem["Created"] = $createdDate;
                $spFileItem["Modified"] = $modifiedDate;
                $spFileItem.UpdateOverwriteVersion();
                return $goodLogin; # file properties updated
            } else {
                return $null; # no properties updated
            };

        }; # function FixAuthor




        function ReportDocLibFolder($spFolder) {
            Write-Verbose "ReportDocLibFolder invoked";
            Write-Verbose "spFolder.Name: $($spFolder.Name)";
            $spFiles = $spFolder.Files;
            $folderOk = $true;
            foreach ($spFile in $spFiles) {
                Write-Verbose "spFile.Name: $($spFile.Name)";
                $spFileItem = $spFile.Item;
                
                
                if ($spFileItem.Properties -eq $null)
                {
                    WriteOutput "ERROR" "ITEM_PROPERTIES_NULL" "subfolder: $($spFolder.Name)`tfile: $($spfile.Name)`tspFileItem.Properties is null, skipping this item";
                    continue;
                };

                $author = $spFileItem.Properties["vti_author"];
                $modifiedby = $spFileItem.Properties["vti_modifiedby"];
                $spUser = $spWeb.SiteUsers[$author];

                if ($spUser -eq $null) {
                    WriteOutput "WARNING" "INVALID_AUTHOR" "subfolder: $($spFolder.Name)`tfile: $($spfile.Name)`tauthor: $author`tmodifiedby: $modifiedby";
                    $folderOk = $false;
                    if ($Script:invalidAuthorLogins -notcontains $author) {
                        $Script:invalidAuthorLogins += $author;
                    };
                } else {
                    WriteOutput "INFO" "FOLDER_OK" "subfolder: $($spFolder.Name)`tfile: $($spfile.Name)`tauthor: $author`tmodifiedby: $modifiedby";
                };
            };
            if ($folderOk) {
                $Script:goodFolderCount++;
            } else {
                $Script:badFolderCount++;
            };
        }; # ReportDocLibFolder


        function FixDocLibFolder($spFolder) {
            Write-Verbose "FixDocLibFolder invoked";
            $spFiles = $spFolder.Files;
            foreach ($spFile in $spFiles) {
                $spFileItem = $spFile.Item;

                if ($spFileItem.Properties -eq $null)
                {
                    WriteOutput "ERROR" "ITEM_PROPERTIES_NULL" "subfolder: $($spFolder.Name)`tfile: $($spfile.Name)`tspFileItem.Properties is null, skipping this item";
                    continue;
                };

                $author = $spFileItem.Properties["vti_author"];
                $spUser = $spWeb.SiteUsers[$author];

                if ($spUser -eq $null) {
                    WriteOutput "WARNING" "INVALID_AUTHOR" "subfolder: $($spFolder.Name)`tfile: $($spfile.Name)`tauthor: $author`tmodifiedby: $modifiedby";

                    $newAuthor = FixAuthor $spFile $author;
                    if ($newAuthor -eq $null) {
                        WriteOutput "ERROR" "NO_MAPPING" "subfolder: $($spFolder.Name)`tfile: $($spfile.Name)`tauthor: $author`tLoginMappings hashtable does not contain an entry for this invalid login, skipping this item";
                        $Script:unfixedFolderCount++;
                    } else {
                        WriteOutput "INFO" "FILE_FIXED" "subfolder: $($spFolder.Name)`tfile: $($spfile.Name)`tauthor: $author`tmodifiedby: $modifiedby`tnewAuthor: $newAuthor";
                        $Script:fixedFolderCount++;
                    };

                };
            };
        }; # function FixDocLibFolder

}