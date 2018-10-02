Function Import-Pictures {
<#
.SYNOPSIS
    This function imports and orders pictures from a memory card into computer folder

.DESCRIPTION
    This function imports pictures and other acceptable files from the current working folder
    (usually a memory card), and imports them in a computer folder.
    Onte way, the filenames are timestamped, i.e. the date and time of creation of the file are used to 
    prefix the file name.
    Additionnaly, the files are arranged under the chosen target folder in a directory tree similar to
        ...\yyyy\yyyymm\yyyymmdd
    If such a directory does not exist, it is created
    If such a directory does exist, or a similar directory with a name or a description suffix after the day date, this directory is used


.PARAMETER Command
    Indicates which action is to be performed:
        Copy: the pictures are copied from the source to the target
        Move: the pictures are moved from the source to the target, i.e. they are deleted from the source
        Offset: the timestamps included in the names of the files in the the target folder are offset by a provided nmber of hours

.PARAMETER DryRun
    The action is not actually performed

.PARAMETER Force
    Existing files on the target folder are overwritten (without further confirmation)

.PARAMETER TargetFolder
    The root folder targeted by the import

.PARAMETER ExcludeTargetFolder
    These folders are not to be used as target, even if the day date matches.

.PARAMETER SubFolder
    If provided, this value is used as an additional directory level under the year level.

.PARAMETER Suffix
    If provided, this value is added to the file names (for instance to identify the camera used to shoot the pictures)

.PARAMETER MinDate
    The files created before this date are not imported

.PARAMETER MaxDate
    The files created before after date are not imported

.PARAMETER Offset
    The creation date of the files is offset by this number of hours before timestamping their names

.EXAMPLE
    ...
    
.NOTES
    Author: Philippe Raemy
    Last Edit: 2018-27-09
    Version 1.0 - initial release


#>
    [CmdletBinding(PositionalBinding=$false)]  # Add cmdlet features.
    Param (
        # Define parameters below, each separated by a comma

        [Parameter(Mandatory=$True)]
        [ValidateSet('Copy', 'Move', 'Offset')]
        [string]$Command,

        [Parameter(Mandatory=$False)]
        [switch]$DryRun,

        [Parameter(Mandatory=$False)]
        [switch]$Force,

        [Parameter(Mandatory=$False)]
        [string]$TargetFolder = 'd:\users\public\pictures',

        [Parameter(Mandatory=$False)]
        [string[]]$ExcludeTargetFolder = '',

        [Parameter(Mandatory=$False)]
        [string[]]$Filter = ('*.jpg', '*.jpeg', '*.mov', '*.mp?'),

        [Parameter(Mandatory=$False)]
        [string]$SubFolder = '',

        [Parameter(Mandatory=$False)]
        [string]$Suffix = '',

        [Parameter(Mandatory=$False)]
        [DateTime]$MinDate = (New-Object System.DateTime(1900,1,1)),

        [Parameter(Mandatory=$False)]
        [DateTime]$MaxDate = (New-Object System.DateTime(2500,1,1)),

        [Parameter(Mandatory=$False)]
        [int]$Offsethours = 0

    )

    Begin {
        # Start of the BEGIN block.
        Write-Verbose -Message "Starting [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)] with parameters:"
        Write-Verbose (@{
            Command             = $Command            ;
            DryRun              = $DryRun             ;
            Force               = $Force              ;
            TargetFolder        = $TargetFolder       ;
            ExcludeTargetFolder = $ExcludeTargetFolder;
            Filter              = $Filter             ;
            SubFolder           = $SubFolder          ;
            Suffix              = $Suffix             ;
            MinDate             = $MinDate            ;
            MaxDate             = $MaxDate            ;
            Offsethours         = $Offsethours
        } | Out-String)

    } # End Begin block

    Process {
        Function Convert-Statistics {
        param(
            [parameter(ValueFromPipeline=$true)]
            [System.IO.FileInfo]
            $file
        )
            Begin{
                $countFiles = 0
                $totalSize = 0
            }
 
            PROCESS
            {
                echo @{File = $file; TotalSize = $totalSize + $file.Length}
            }
        }


        dir $Filter -Recurse `
            | Where-Object -Property CreationTime -GE $MinDate `
            | Where-Object -Property CreationTime -LE $MaxDate `
            | Convert-Statistics `
            | Format-Table
            # | Where-Object -FilterScript {echo $_}

    } # End of PROCESS block.

    End {
        # Start of END block.
        Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."
 
        # Add additional code here.
 
    } # End of the END Block.

}

