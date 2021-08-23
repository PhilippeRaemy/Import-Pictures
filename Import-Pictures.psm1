Function Import-Pictures {
<#
.SYNOPSIS
    This function imports and orders pictures from a memory card into 
    a local folder

.DESCRIPTION
    This function imports pictures and other acceptable files from the 
    current working folder (usually a memory card), and imports them in 
    a local folder.
    On the way, the filenames are timestamped, i.e. the date and time of 
    creation of the file are used to prefix the file name, inan ISO format.
    Additionnaly, the files are arranged under the chosen target folder 
    in a directory tree similar to
        ...\yyyy\yyyymm\yyyymmdd
    If such a directory does not exist, it is created.
    If such a directory does exist, or a similar directory with a name or 
    a description suffix after the day date, this directory is used.


.PARAMETER Command
    Indicates which action is to be performed:
        Copy: the pictures are copied from the source to the target.
        Move: the pictures are moved from the source to the target, i.e. 
            they are deleted from the source.
        Offset: the timestamps included in the names of the files in 
            the target folder are offset by a provided nmber of hours.

.PARAMETER DryRun
    The action is not actually performed

.PARAMETER Force
    Any existing file on the target folder will be overwritten (without further 
    confirmation)

.PARAMETER TargetFolder
    The root folder targeted by the import.

.PARAMETER ExcludeTargetFolder
    These folders are not to be used as target, even if the day date matches.

.PARAMETER SubFolder
    If provided, this value is used as an additional directory level under 
    the year level.

.PARAMETER Suffix
    If provided, this value is added to the file names (for instance to 
    identify the camera used to shoot the pictures)

.PARAMETER MinDate
    The files created before this date are not imported

.PARAMETER MaxDate
    The files created before after date are not imported

.PARAMETER Offset
    The creation date of the files is offset by this number of hours before 
    timestamping their names

.EXAMPLE
    ...

.NOTES
    Author: Philippe Raemy
    Last Edit: 2018-11-02
    Version 0.1 - initial release


#>
    [CmdletBinding(PositionalBinding=$false)]  # Add cmdlet features.
    Param (
        [Parameter(Mandatory=$True )] [ValidateSet('Copy', 'Move', 'Offset')] [string]$Command,
        [Parameter(Mandatory=$False)] [switch]   $DryRun,
        [Parameter(Mandatory=$False)] [switch]   $Force,
        [Parameter(Mandatory=$False)] [string]   $TargetFolder = '',
        [Parameter(Mandatory=$False)] [string[]] $ExcludeTargetFolder,
        [Parameter(Mandatory=$False)] [string[]] $Filter       = ('*.jpg', '*.jpeg', '*.mov', '*.mp?', '*.cr2'),
        [Parameter(Mandatory=$False)] [string]   $SubFolder    = '',
        [Parameter(Mandatory=$False)] [string]   $Suffix       = '',
        [Parameter(Mandatory=$False)] [DateTime] $MinDate      = (New-Object System.DateTime(1900,1,1)),
        [Parameter(Mandatory=$False)] [DateTime] $MaxDate      = (New-Object System.DateTime(2500,1,1)),
        [Parameter(Mandatory=$False)] [int]      $Offsethours  = 0

    )

    Begin {
        # Start of the BEGIN block.
        if($TargetFolder -eq '') {
            if($env:computername -eq 'ZENBOOK') {$TargetFolder = 'C:\Users\Philippe\OneDrive - RL&Kids\Pictures\'}
            elseif($env:computername -eq 'SERVER02') {$TargetFolder = 'd:\users\public\pictures'}
            else {$TargetFolder = 'c:\users\public\pictures'}
        }

        Write-Verbose -Message "Starting [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)] on $($env:computername) with Parameters:"
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
        ######################################################################################
        Function New-FileDetails{
            param(
                [Parameter(ValueFromPipeline=$True)] [System.IO.FileInfo] $file
            )
            BEGIN
            {
                $countFiles = 0
                $totalSize = [int64]0
            }

            PROCESS
            {
                $totalSize += $file.Length
                $countFiles++;
                Write-Output @{
                    TotalSize      = $totalSize;
                    Length         = $file.Length;
                    CreationTime   = $file.Length;
                    Position       = $countFiles;
                    File           = $file;
                }
                Write-Verbose "$file, $countFiles"
            }
        }
        ######################################################################################
        Function Format-Output{
            param(
                [Parameter(ValueFromPipeline=$True)] $f,
                [Parameter(Mandatory=$True)] [int64] $expectedSize,
                [Parameter(Mandatory=$True)] [int]   $expectedCount,
                [Parameter(Mandatory=$True)] [string]$Activity
            )
            BEGIN
            {
                $expectedSizeMb = [Math]::Round($expectedSize / 1048576, 1);
                $progress = "Progress {position}/$expectedCount ({progress}%)" + `
                    " - {totalSize}Mb/$($expectedSizeMb)Mb ({progressMb}%)";
            }
            PROCESS
            {
               $progressMb = 100*$f.TotalSize/$expectedSize;
               $status = $progress `
                    -replace '{position}'  , $f.Position `
                    -replace '{totalSize}' , ('{0:N1}' -f [Math]::Round($f.TotalSize/1048576, 1)) `
                    -replace '{progress}'  , ('{0:N2}' -f [Math]::Round(100*$f.Position/$expectedCount, 2)) `
                    -replace '{progressMb}', ('{0:N1}' -f [Math]::Round($progressMb, 1))
               Write-Progress -Activity $activity -PercentComplete $progressMb -Status $status
               return $f.Message
            }
        }
        ######################################################################################
        Function Invoke-Action{
            param(
                [Parameter(ValueFromPipeline=$True)]   $f,
                [Parameter(Mandatory=$True)] [string] $Command,
                [Parameter(Mandatory=$True)] [bool]   $DryRun,
                [Parameter(Mandatory=$True)] [bool]   $Force
            )

            PROCESS
            {
                try{
                    $doIt = $Force -or -not (Test-Path -Path $f.Location)
                    $verb = if($DryRun) {'would be'} else {'is'}
                    if($Command -eq 'Copy'){
                        if($doIt) {
                            $f.Message = "$($f.File) $verb copied to $($f.Location)."
                            if(-not $DryRun) {
                                $f.Location.Directory.Create()
                                copy $f.File $f.Location.FullName
                            }
                        } else {
                            $f.Message = "$($f.File) exists as $($f.Location)."
                        }
                    }
                    elseif($Command -eq 'Move'){
                        if($doIt) {
                            $f.Message = "$($f.File) $verb moved to $($f.Location).";
                            if(-not $DryRun) {
                                $f.Location.Directory.Create()
                                move $f.File $f.Location -Force
                            }
                        } else {
                            if(-not $DryRun) {del $f.File}
                            $f.Message = "$($f.File) $verb deleted";
                        }
                    }
                    elseif($Command -eq 'Offset'){
                        $f.Message = "$($f.File) $verb renamed to $($f.Location.Name).";
                        if(-not $DryRun) {ren $f.File $f.Location.Name -Force  }
                    }
                }
                catch
                {
                    $f.Message = "$($_.Exception.GetType().FullName): $($_.Exception.Message) while trying: $($f.Message)."
                }
                return $f
            }
        }
        ######################################################################################
        Function Where-NotExcluded{
            param(
                [Parameter(ValueFromPipeline=$True)] $f,
                [Parameter(Mandatory=$True)] [AllowNull()] [string[]]$ExcludeTargetFolder
            )

            if($ExcludeTargetFolder){
                $split = $_.Split([System.IO.Path]::DirectorySeparatorChar)
                if($ExcludeTargetFolder | %{$split -match $_ }){
                    return;
                }
            }
            return $f;
        }
        ######################################################################################
        Function Resolve-Location{
            param(
                [Parameter(ValueFromPipeline=$True)] $f,
                [Parameter(Mandatory=$True)]                      [string]   $TargetFolder,
                [Parameter(Mandatory=$True)] [AllowEmptyString()] [string]   $SubFolder,
                [Parameter(Mandatory=$True)] [AllowNull()]        [string[]] $ExcludeTargetFolder,
                [Parameter(Mandatory=$True)]                      [int]      $Offsethours
            )
            PROCESS
            {
                Write-Verbose "Resolve-Location: $($f.file) $($f.Position)"
                $creationTime = $f.file.CreationTime
                [ref] $creationTimeRef = $creationTime
                $fileIsDated = $false
                $DateInName = $f.file.Name -match '(\d{8})[^\d](\d{4,6})'
                if($DateInName){
                    $fileIsDated = [DateTime]::TryParseExact("$($matches[1])_$($matches[2])", `
                        'yyyyMMdd_HHmmss', `
                        [System.Globalization.CultureInfo]::InvariantCulture, `
                        [System.Globalization.DateTimeStyles]::None,`
                        $creationTimeRef)
                    if($fileIsDated) {
                        $creationTime = $creationTimeRef.Value
                    }
                } else {
                    $DateInName = $f.file.Name -match '(\d{8})'
                    if($DateInName){
                        $fileIsDated = [DateTime]::TryParseExact($matches[1], `
                            'yyyyMMdd', `
                            [System.Globalization.CultureInfo]::InvariantCulture, `
                            [System.Globalization.DateTimeStyles]::None,`
                            $creationTimeRef)
                        if($fileIsDated) {
                            $creationTime = $creationTimeRef.Value
                        }
                    }
                }

                $creationTime = $creationTime.AddHours($Offsethours)
                
                $filename = if($fileIsDated) {$f.file.Name.Replace($matches[0], '')} else {$f.file.Name}
                $separator = if($filename.Substring(1, 0) -ne '_') {'_'} else {''}
                $filename = $creationTime.ToString("yyyyMMdd_HHmmss") + $separator + $filename

                $folderRoot = [System.IO.Path]::Combine($TargetFolder, $creationTime.ToString('yyyy'), $SubFolder)
                if(Test-Path -Path $folderRoot){
                    $folder = (dir $folderRoot -Directory -Recurse $creationTime.ToString('yyyyMMdd*') -ErrorAction SilentlyContinue `
                        | Where-NotExcluded -ExcludeTargetFolder $ExcludeTargetFolder `
                        | select -First 1).FullName
                }
                else {
                    $folder = $null
                }

                Write-Verbose "folderRoot is $folderRoot"
                Write-Verbose "folder is $folder"
                if(-not $folder){
                    $folder = [System.IO.Path]::Combine($folderRoot, $creationTime.ToString('yyyyMM'), $creationTime.ToString('yyyyMMdd'))
                    Write-Verbose "new folder is $folder"
                }

                $f.Location = New-Object System.IO.FileInfo([System.IO.Path]::Combine($folder, $filename))
                Write-Output $f
            }
        }
        ######################################################################################
        $workAtHand = dir $Filter -Recurse `
            | Where-Object -Property CreationTime -GE $MinDate `
            | Where-Object -Property CreationTime -LE $MaxDate

        $totalSize = $workAtHand | Measure -Property Length -Sum

        $workAtHand `
            | New-FileDetails `            | Resolve-Location -TargetFolder $TargetFolder -SubFolder $SubFolder -ExcludeTargetFolder $ExcludeTargetFolder -Offsethours $Offsethours `            | Invoke-Action    -Command $Command -DryRun $DryRun.IsPresent -Force $Force.IsPresent `            | Format-Output    -ExpectedSize $totalSize.Sum -ExpectedCount $totalSize.Count -Activity "Import-Pictures $(if($DryRun) {'Try '})$Command..." `
            | Format-Table

    } # End of PROCESS block.

    End {
        # Start of END block.
        # Write-Verbose -Message "Entering the END block [$($MyInvocation.MyCommand.CommandType): $($MyInvocation.MyCommand.Name)]."

        # Add additional code here.

    } # End of the END Block.

}