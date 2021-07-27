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
# Last change: 27.07.2021
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

# Extract prerequisites
Write-Host "Extracting prerequisites, please wait..."
Expand-Archive -Path "$PSScriptRoot\_Prerequisites.zip" -DestinationPath $tempPath -Force -ErrorAction Stop

$sourceDirectory = Read-Host -Prompt "Enter the source folder path. It will be scanned for all JPEG files"
$sourceDirectory = $sourceDirectory.Replace('"','')

Write-Host "Enumerating files, please wait..."
$process = Start-Process -FilePath $tempPath\exiftool.exe `
                         -ArgumentList @(# Check for Samsung:EmbeddedVideoFile
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
                                         # Check for Google Micro Video Offset
                                         "-if",
                                         '"defined $xmp:MicroVideoOffset"',
                                         "-p",
                                         '"Google.MicroVideo|$MicroVideoOffset|$directory\$filename"',
                                         # Common arguments
                                         "-common_args"
                                         "-r",
                                         "-ext",
                                         "jpg",
                                         "-ext",
                                         "jpeg",
                                         """$sourceDirectory""") `
                         -RedirectStandardError $tempPath\exifPS-StdErr.txt `
                         -RedirectStandardOutput $tempPath\exifPS-StdOut.txt `
                         -PassThru `
                         -WindowStyle Hidden

while($process.HasExited -eq $false)
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
                $argumentList = @("-a",
                                  "-b",
                                  "-$($file.Type.Replace('.',':'))",
                                  "-charset",
                                  "filename=latin1",
                                  """$($fileObject.FullName)"""
                                 )
                Write-Host "$argumentList" -ForegroundColor Cyan
                $process = Start-Process -FilePath $tempPath\exiftool.exe `
                                         -ArgumentList $argumentList `
                                         -Wait `                                         -PassThru `
                                         -RedirectStandardError $tempPath\exifPS-StdErr.txt `
                                         -RedirectStandardOutput $videoTempFilePath `
                                         -ErrorAction Stop `
                                         -WindowStyle Hidden

                if ($process.ExitCode -ne 0)
                    {
                        Write-Host "      The following error has occured:" -ForegroundColor Red
                        Get-Content $tempPath\exifPS-StdErr.txt -Encoding UTF8
                        Read-Host "Press any key to stop this script"
                        exit
                    }

                Remove-Item $tempPath\exifPS-StdErr.txt -Force
            }
        elseif ($file.Type -eq "Google.MicroVideo")
            {
                Write-Host "GoogleMicroVideo from offset $($file.MicroVideoOffset)" -ForegroundColor Cyan
                Get-Content $fileObject.FullName -Raw -Encoding Byte | % { $_[($_.Length-$($file.MicroVideoOffset)) ..($_.Length-1)] } | Set-Content $videoTempFilePath -Encoding Byte
            }

        # Encode the video
        $videoFilePath = "$($fileObject.DirectoryName)\$($fileObject.BaseName).$($file.Type).mp4"
        $argumentList = @("-i",
                        """$videoTempFilePath""",
                        "-c:v",
                        "libx265",
                        "-x265-params",
                        "deblock=4,4",
                        "-crf",
                        "35",
                        "-preset",
                        "slower",
                        """$videoFilePath"""
                       )
        Write-Host "   Encoding video: " -NoNewline
        Write-Host "$tempPath\ffmpeg.exe $argumentList" -ForegroundColor Cyan
        Start-Process -FilePath powershell `
                      -ArgumentList @("-ExecutionPolicy",
                                      "Bypass",
                                      "Write-Host 'Setting ffmpeg priority...'; Start-Sleep 3; (Get-Process -Name ffmpeg).PriorityClass = 'Idle'; Start-Sleep 3"
                                     ) `
                      -WindowStyle Hidden
        $process = Start-Process -FilePath "$tempPath\ffmpeg.exe" `
                                 -ArgumentList $argumentList `                                 -Wait `
                                 -PassThru `
                                 -RedirectStandardError "$tempPath\ffmpeg.log" `
                                 -ErrorAction Stop `
                                 -WindowStyle Hidden

        if ($process.ExitCode -ne 0)
            {
                Write-Host "      The following error has occured:" -ForegroundColor Red
                Get-Content $tempPath\ffmpeg.log -Encoding UTF8
                Read-Host "Press any key to stop this script"
                exit
            }

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
                $content | Set-Content $fileObject.FullName -Encoding Byte
            }
        Write-Host "   Removing video data and unnecessary EXIF info from source file: " -NoNewline
                $argumentList = @("-xmp:MicroVideo=",
                                  "-xmp:MicroVideoOffset=",
                                  "-xmp:MicroVideoPresentationTimestampUs=",
                                  "-xmp:MicroVideoVersion=",
                                  "-trailer:all=",
                                  "-charset",
                                  "filename=latin1",
                                  "-overwrite_original",
                                  """$($fileObject.FullName)"""
                                 )
                
                Write-Host "exiftool $argumentList" -ForegroundColor Cyan
                $process = Start-Process -FilePath $tempPath\exiftool.exe `
                                         -ArgumentList $argumentList `
                                         -Wait `
                                         -PassThru `
                                         -RedirectStandardError $tempPath\exifPS-StdErr.txt `
                                         -RedirectStandardOutput $tempPath\exifPS-StdOut.txt `
                                         -ErrorAction Stop `
                                         -WindowStyle Hidden

                if ($process.ExitCode -ne 0)
                    {
                        Write-Host "      The following error has occured:" -ForegroundColor Red
                        Get-Content $tempPath\exifPS-StdErr.txt -Encoding UTF8
                        Read-Host "Press any key to stop this script"
                        exit
                    }

        # Copy EXIF tags over
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
        Write-Host "   Copying EXIF tags from source to target file: " -NoNewline
        Write-Host "$argumentList" -ForegroundColor Cyan
        $process = Start-Process -FilePath $tempPath\exiftool.exe `
                                 -ArgumentList $argumentList `
                                 -Wait `
                                 -PassThru `
                                 -RedirectStandardError $tempPath\exifPS-StdErr.txt `
                                 -RedirectStandardOutput $tempPath\exifPS-StdOut.txt `
                                 -ErrorAction Stop `
                                 -WindowStyle Hidden

        if ($process.ExitCode -ne 0)
            {
                Write-Host "      The following error has occured:" -ForegroundColor Red
                Get-Content $tempPath\exifPS-StdErr.txt -Encoding UTF8
                Read-Host "Press any key to stop this script"
                exit
            }

        $i++
    }

# Cleanup
Write-Host "All done. Please check if all went right and then press Enter to finish up."
Read-Host
Write-Host "Removing temp files..."
Remove-Item $tempPath -Recurse -Force