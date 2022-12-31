###################################################################
#
# Script to split video and image from various motion photo formats
# e.g. Samsung Surround Shot Video, Google MVIMG
#
# - Tested with exiftool 12.28
# - Tested with ffmpeg 4.4-full_build-www.gyan.dev
#
# _Prerequisites.zip must be present in the script's root folder
# and consist of exiftool.exe and ffmpeg.exe
#
# Author: Mark Hermann
# Last change: 31.12.2022
# 30.12.2022: Added functionality for newer Google Motion Photos
#
###################################################################

# Variables
# ---------
$tempPath = "$($env:TEMP)\$($MyInvocation.MyCommand.Name)"
$tagsToCopy = @("-CreateDate",         # Which tags should be copied from source JPG to video file
                "-DateTimeOriginal",
                "-TrackCreateDate",
                "-MediaCreateDate",
                "-GPSCoordinates",
                "-GPSLatitude",
                "-GPSLongitude",
                "-GPSPosition")
$exifToolPath = "D:\Nextcloud\Eigene Dateien\Programmieren & Basteln\Exiftool Skripte\exiftool.exe"
$ffmpegPath   = "D:\Nextcloud\Eigene Dateien\Programmieren & Basteln\FFmpeg & Skripte\FFQueue_1_7_58\ffmpeg.exe"

# Functions
# ---------
# Needed function to find trailing rubbish (see fixes section below)
function Find-Bytes ([byte[]]$Bytes, [byte[]]$Search, [int]$Start, [Switch]$All)
    {
        for ($Index = $Start; $Index -le $Bytes.Length - $Search.Length ; $Index++)
            {
                for ($i = 0; $i -lt $Search.Length -and $Bytes[$Index + $i] -eq $Search[$i]; $i++) {}
                if ($i -ge $Search.Length)
                    { 
                        $Index
                        if (!$All)
                            {
                                Return
                            }
                    } 
            }
    }

function Run-Command ([String]$commandName, $argumentList, [String]$stdOutPath="$tempPath\exifPS-StdOut.txt", [Switch]$wait=$false)
    {
        $commandName = """$commandName"""
        Write-Host "$commandName $argumentList" -ForegroundColor Cyan
        New-Variable -Name process -Value $null -Scope Global -Force
        $global:process = Start-Process -FilePath "$commandName" `
                                 -ArgumentList $argumentList `
                                 -Wait:$wait `                                 -PassThru `
                                 -RedirectStandardError $tempPath\exifPS-StdErr.txt `
                                 -RedirectStandardOutput $stdOutPath `
                                 -ErrorAction Stop `
                                 -WindowStyle Hidden

        if ($wait -eq $true)
            {
                if ($global:process.ExitCode -ne 0)
                    {
                        Write-Host "      The following error has occured:" -ForegroundColor Red
                        Get-Content $tempPath\exifPS-StdErr.txt -Encoding UTF8
                        Read-Host "Press any key to continue or STRG+C to stop"
                    }

                 Remove-Item $tempPath\exifPS-StdErr.txt -Force
            }
    }

# Extract prerequisites
#Write-Host "Extracting prerequisites, please wait..."
#Expand-Archive -Path "$PSScriptRoot\_Prerequisites.zip" -DestinationPath $tempPath -Force -ErrorAction Stop

if (-Not (Test-Path $tempPath))
    {
        New-Item $tempPath -ItemType Directory | Out-Null
    }

$sourceDirectory = Read-Host -Prompt "Enter the source folder path. It will be scanned for all JPEG files"
$sourceDirectory = $sourceDirectory.Replace('"','')

Write-Host "Enumerating files, please wait..."
Run-Command -commandName $exifToolPath `
            -argumentList @(# Check for Samsung:EmbeddedVideoFile
                                         "-if",
                                         '"defined $Samsung:EmbeddedVideoFile"',
                                         "-p"
                                         '"Samsung.EmbeddedVideoFile|null|$directory\$filename"',
                                         "-execute",
                                         # Check for Samsung:SurroundShotVideo
                                         "-if",
                                         '"defined $Samsung:SurroundShotVideo"',
                                         "-p",
                                         '"Samsung.SurroundShotVideo|null|$directory\$filename"',
                                         "-execute"
                                         # Check for Google Micro Video Offset (old Motion Photos)
                                         "-if",
                                         '"defined $xmp:MicroVideoOffset"',
                                         "-p",
                                         '"Google.MicroVideo|$MicroVideoOffset|$directory\$filename"',
                                         "-execute",
                                         # Check for Google Micro Video Offset (new Motion Photos)
                                         "-if",
                                         '"defined $xmp:MotionPhotoVersion"',
                                         "-p",
                                         '"Google.MotionVideo|null|$directory\$filename"',
                                         # Common arguments
                                         "-common_args"
                                         "-r",
                                         "-ext",
                                         "jpg",
                                         "-ext",
                                         "jpeg",
                                         """$sourceDirectory""")

while($global:process.HasExited -eq $false)
    {
        Write-Host "$(Get-Date): Still enumerating files, please wait... Files found: " -NoNewLine
        Write-Host ((Get-Content $tempPath\exifPS-StdOut.txt) | Measure-Object).Count
        Start-Sleep "5"
    }

# Read files from output file
$filesContent = Get-Content $tempPath\exifPS-StdOut.txt -Encoding UTF8
$filesContent = $filesContent.Replace("/","\")
# Get all unique files that have only one video file reference within the EXIF data
$files = ($filesContent | ConvertFrom-Csv -Delimiter "|" -Header @("Type", "MicroVideoOffset", "Path")  | Sort-Object -Property Path,Type | Group-Object -Property Path | Where {$_.Count -eq 1}).Group
# Get all files that have multiple video file references within the EXIF data and select the one that is faster to extract
$files += ($filesContent | ConvertFrom-Csv -Delimiter "|" -Header @("Type", "MicroVideoOffset", "Path")  | Sort-Object -Property Path,Type | Group-Object -Property Path | Where {$_.Count -ge 2}).Group | where{$_.Type -eq "Samsung.EmbeddedVideoFile"}

$fileCount = ($files | Measure-Object).Count

Write-Host "Please check below for errors. This is StdErr of Exiftool (non 0 return value is normal when using ifs):" -ForegroundColor Red
Get-Content $tempPath\exifPS-StdErr.txt -Encoding UTF8

Remove-Item $tempPath\exifPS-StdErr.txt -Force
Remove-Item $tempPath\exifPS-StdOut.txt -Force

$i = 1
foreach ($file in $files)
    {
        $file.Path = [System.Management.Automation.WildcardPattern]::Escape($file.Path) # Escape special characters in file name like [ and ]
        $fileObject = Get-Item $file.Path
        Write-Host "Processing file $i of $($fileCount): $fileObject"
        
        # Create backup
        $backupFilePath = "$($fileObject.DirectoryName)\$($fileObject.BaseName).SplitOriginal$($fileObject.Extension)"
        Copy-Item $fileObject -Destination $backupFilePath -ErrorAction Stop -WarningAction Stop

        # Extract the video
        $videoTempFilePath = "$tempPath\$($fileObject.BaseName).$($file.Type)Temp.mp4"
        Write-Host "   Extracting video: " -NoNewline
        if ($file.Type -eq "Samsung.SurroundShotVideo" -or $file.Type -eq "Samsung.EmbeddedVideoFile")
            {
                Run-Command -commandName $exifToolPath `
                            -argumentList @("-a",
                                            "-b",
                                            "-$($file.Type.Replace('.',':'))",
                                            "-charset",
                                            "filename=latin1",
                                            """$($fileObject.FullName)"""
                                           ) `
                            -stdOutPath $videoTempFilePath `
                            -wait
            }
        elseif ($file.Type -eq "Google.MicroVideo")
            {
                Write-Host "GoogleMicroVideo from offset $($file.MicroVideoOffset)" -ForegroundColor Cyan
                Get-Content $fileObject.FullName -Raw -Encoding Byte | % { $_[($_.Length-$($file.MicroVideoOffset)) ..($_.Length-1)] } | Set-Content $videoTempFilePath -Encoding Byte
            }
        elseif ($file.Type -eq "Google.MotionVideo")
            {
                $videoTempFilePath
                $searchBytes = @([byte]0x00,0x00,0x00,0x1C,0X66,0x74,0x79,0x70,0x69,0x73,0x6F,0x6D) # [...]ftypisom
                $fileBytes = Get-Content $fileObject.FullName -Raw -Encoding Byte
                $googleMotionPhotoOffset = Find-Bytes -Bytes $fileBytes -Search $searchBytes
                Write-Host "Extracting Google.MotionVideo from offset $googleMotionPhotoOffset" -ForegroundColor Cyan -NoNewline
                #$fileBytes | % { $_[($_.Length-$googleMotionPhotoOffset) ..($_.Length-1)] } | Set-Content $videoTempFilePath -Encoding Byte
                $fileBytes[$googleMotionPhotoOffset ..($fileBytes.Length-1)] | Set-Content $videoTempFilePath -Encoding Byte
                Write-Host " done"
            }

        # Encode the video
        $videoFilePath = "$($fileObject.DirectoryName)\$($fileObject.BaseName).$($file.Type).mp4"
        Write-Host "   Encoding video: " -NoNewline
        Start-Process -FilePath powershell `
                      -ArgumentList @("-ExecutionPolicy",
                                      "Bypass",
                                      "Write-Host 'Setting ffmpeg priority...'; Start-Sleep 3; (Get-Process -Name ffmpeg).PriorityClass = 'Idle'; Start-Sleep 3"
                                     ) `
                      -WindowStyle Hidden
        Run-Command -commandName $ffmpegPath `
                    -argumentList @("-i",
                                    """$videoTempFilePath""",
                                    "-c:v",
                                    #"libx265", #HEVC
                                    "libsvtav1"
                                    #"-x265-params", #HEVC
                                    #"deblock=4,4", #HEVC
                                    "-crf",
                                    "35",
                                    #"-preset", #HEVC
                                    "-preset:v",
                                    #"slower", #HEVC
                                    "4",
                                    "-pix_fmt",
                                    "yuv420p10le",
                                    "-svtav1-params",
                                    "input-depth=10:keyint=10s",
                                    "-map",
                                    "0:v:0",
                                    """$videoFilePath"""
                                   ) `
                    -wait

        $sourceSize = (Get-Item $videoTempFilePath).Length
        $targetSize = (Get-Item $videoFilePath).Length
        $spaceSaved = $sourceSize - $targetSize
        if ((100/$sourceSize*$targetSize) -gt 70)
            {
                Write-Host "   Space saved $([math]::Round($spaceSaved/ 1024 / 1024, 2))mb is less than 30%" -ForegroundColor Red
                Read-Host "Press Enter to continue"
            }
        else
            {
                Write-Host "   Space saved $([math]::Round($spaceSaved/ 1024 / 1024, 2))mb" -ForegroundColor Green
            }

        # Remove the video bitstream from the original
        if ($file.Type -eq "Google.MicroVideo")
            {
                Write-Host "   Removing video data and unnecessary EXIF info from source file: " -NoNewline
                Write-Host "GoogleMicroVideo from offset $($file.MicroVideoOffset) to end of file" -ForegroundColor Cyan
                $content = Get-Content $fileObject.FullName -Raw -Encoding Byte | % { $_[0 ..($_.Length-$($file.MicroVideoOffset))] }
                Set-Content $fileObject.FullName -Encoding Byte -Value $content
            }
        Write-Host "   Removing video data and unnecessary EXIF info from source file: " -NoNewline
        Run-Command -commandName $exifToolPath `
                    -argumentList @("-xmp:MicroVideo=",
                                    "-xmp:MicroVideoOffset=",
                                    "-xmp:MicroVideoPresentationTimestampUs=",
                                    "-xmp:MicroVideoVersion=",
                                    "-xmp:MotionPhoto=",
                                    "-xmp:MotionPhotoVersion=",
                                    "-xmp:MotionPhotoPresentationTimestampUs=",
                                    "-trailer:all=",
                                    "-charset",
                                    "filename=latin1",
                                    "-overwrite_original",
                                    """$($fileObject.FullName)"""
                                    ) `
                    -wait

        #################################################################################################################################
        # Fix various issues that arise within the newly created files
        # This is primarily due to Nextcloud being unable to read those files (it wasn't able to read those before splitting them either)
        # e.g. no EOI marker is found or sometimes its at the wrong position
        # Searching online this seems to be a Samsung bug that hasn't been fixed in years

        if ($file.Type -eq "Samsung.SurroundShotVideo")
            {
                # 1. First of all we will remove data rubbish that is left within the file after the actual bitstream of compressed JPEG data
                # This is identified by a specific hex string
                # 00 00 01 02 17 00 00 00 4D 6F 74 69 6F 6E 5F 50 61 6E 6F 72 61 6D 61 = "[Unreadable Data]Motion_Panorama"
                $content = Get-Content $fileObject.FullName -Raw -Encoding Byte
                $searchBytes = [BYTE[]]@(0x00, 0x00, 0x01, 0x02, 0x17, 0x00, 0x00, 0x00, 0x4D, 0x6F, 0x74, 0x69, 0x6F, 0x6E, 0x5F, 0x50, 0x61, 0x6E, 0x6F, 0x72, 0x61, 0x6D, 0x61)
                $eoiMarker = [BYTE[]]@(0xFF, 0xD9)
                $indexOfTrailingRubbish = Find-Bytes -Bytes $content -Search $searchBytes -All
                if ($indexOfTrailingRubbish -ne $null)
                    {
                        $content  = $content[0..($indexOfTrailingRubbish-1)]
                        $content += $eoiMarker
                        Set-Content $fileObject.FullName -Encoding Byte -Value $content
                    }
                # End of 1. fix

                # 2. Now we will remove broken IFD1 tags as Samsung doesn't follow JPEG specifications
                Write-Host "   Removing broken Samsung IFD1 tags from target file: " -NoNewline
                Run-Command -commandName $exifToolPath `
                            -argumentList @("-ifd1:all="
                                          "-overwrite_original",
                                          "-charset",
                                          "filename=latin1",
                                          """$($fileObject.FullName)"""
                                          ) `
                            -wait
            }
        # End of 2. fix

        # 3. Now we will fix possible remaining issues with the EXIF data
        # -ExifVersion=0232 upgrades to EXIF version 2.32
        # Needed as Samsung writes LensModel, but uses wrong EXIF version...
        Write-Host "   Fixing possible EXIF issues within target file: " -NoNewline
        Run-Command -commandName $exifToolPath `
                    -argumentList ("-all=",
                                   "-TagsFromFile",
                                   "@",
                                   "-all:all",
                                   "-unsafe",
                                   "-icc_profile",
                                   "-execute",
                                   "-ExifVersion=0232",
                                   "-execute",
                                   "-common_args",
                                   "-overwrite_original",
                                   "-charset",
                                   "filename=latin1",
                                   """$($fileObject.FullName)"""
                                   ) `
                    -wait
        # Enf of 3. fix


        # End of fixes
        #################################################################################################################################

        # Copy EXIF tags over
        Write-Host "   Copying EXIF tags from source to target video file: " -NoNewline
        $argumentList = @("-TagsFromFile",
                          """$($fileObject.FullName)""",
                          "-charset",
                          "filename=latin1",
                          "-overwrite_original"
                          "-api",
                          "QuickTimeUTC"
                          )
        $argumentList += $tagsToCopy
        $argumentList += @("""$videoFilePath""")
        
        Run-Command -commandName $exifToolPath `
                    -argumentList $argumentList `
                    -wait

        # Final JPEG verification
        Write-Host "   Verifying newly generated JPG: " -NoNewline
        Run-Command -commandName $exifToolPath `
                    -argumentList ("-validate",
                                   "-warning",
                                   "-a",
                                   """$($fileObject.FullName)""",
                                   "-charset",
                                   "filename=latin1"
                                    ) `
                    -wait
        
        $stdOut = Get-Content $tempPath\exifPS-StdOut.txt -Encoding UTF8
        if ($stdOut -ne "Validate                        : OK")
            {
                Write-Host "      The following errors were found or have occured:" -ForegroundColor Red
                Get-Content $tempPath\exifPS-StdErr.txt -Encoding UTF8
                $stdOut
                Read-Host "Press any key to continue or STRG+C to stop this script"
            }

        $i++
    }

# Cleanup
Write-Host "All done. Please check if all went right and then press Enter to finish up."
Read-Host
Write-Host "Removing temp files..."
Remove-Item $tempPath -Recurse -Force