

function Write-Log {
    param(
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$Message,

        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [ValidateSet('Information','Warning','Error')]
        [string]$Severity = 'Information',

        # Name of the resource group that this log entry refers to
        [Parameter()]
        [ValidateNotNullOrEmpty()]
        [string]$LogFileName
    )

    [pscustomobject]@{
        "Time (UTC)" = ([system.datetime]::Utcnow.tostring('u').replace(' ','T'))
        Severity = $Severity
        Message = $Message
    } | Export-Csv -Path  $LogFileName -Append -NoTypeInformation -Delimiter "`t"
}

function Convert-Image {
    param (
        # File name of the image to be converted
        [Parameter(Mandatory)]
        [System.IO.FileInfo]
        $imagefile,
    
        # Text to be printed on the image
        [Parameter(Mandatory)]
        [System.Object]
        $imagetext
    )

    $sFormat = [System.Drawing.StringFormat]::new()
    $sFormat.alignment = [System.Drawing.StringAlignment]::Center
    $sFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
    $font1 = [System.Drawing.Font]::new("Segoe UI",18)
    $font2 = [System.Drawing.Font]::new("Segoe UI",10)
    $textbrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::Black)
    $fillbrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(96,255,255,255))

    $bmp = [System.Drawing.Bitmap]::FromFile($imageFile)
    $image = [System.Drawing.Graphics]::FromImage($bmp)
    $SR = $bmp | Select-Object Width,Height
    $targetImageFileName = $imagefile.DirectoryName + "\" + $imagefile.BaseName + ".bmp"

    $szText = $image.MeasureString(($imageText.Title + " | " + $imageText.Description), $font2)
    $rectText = [System.Drawing.RectangleF]::new((($SR.Width / 2) - ($szText.Width / 2)), ($SR.Height - ($szText.Height * 4)), $szText.Width, $szText.Height)
    $image.FillRectangle($fillbrush, $rectText)
    $image.DrawString(($imageText.Title + " | " + $imageText.Description), $font2, $textbrush, $rectText, $sFormat)
    
    $image.Dispose()
    $bmp.Save($targetImageFileName, [System.Drawing.Imaging.ImageFormat]::Bmp)
    Write-Log -Message "Wrote captioned image file: $targetImageFileName" -LogFileName $logfilepath
    $bmp.Dispose()

    Get-Item $targetImageFileName | Write-Output

}


$wallpaperdir = join-path -Path $env:USERPROFILE -ChildPath "Pictures" -AdditionalChildPath "BingWallpaper"
$logfile = "BingWallpaper.log"


$logfilepath = join-path -path $wallpaperdir -ChildPath $logfile
$imageHistoryFile = Join-Path -Path $wallpaperdir -ChildPath "ImageHistory.csv"

$wallpaperIndex = 0

$imageHistory = @()

$daysToGet = 5

$filenames = @()

if (-not (Test-Path "$wallpaperdir")) {
    New-Item -ItemType Directory -Path "$wallpaperdir"
    Write-Log -Message "Created new wallpaper directory: $wallpaperdir" -LogFileName $logfilepath
}
# Delete old wallpaper files
Get-ChildItem -Path $wallpaperdir -Filter *.bmp | Remove-Item

Write-Log -Message "****************************************" -LogFileName $logfilepath
Write-Log -Message "New run --- $([system.datetime]::Utcnow.tostring('u').replace(' ','T'))" -LogFileName $logfilepath
Write-Log -Message "****************************************" -LogFileName $logfilepath

$desktopCount = Get-DesktopCount

$bingimagedata = Invoke-RestMethod -Uri "https://www.bing.com/HPImageArchive.aspx?format=js&idx=$wallpaperIndex&n=$desktopCount&mkt=en-GB" -Method Get

for ($i = 0; $i -lt $desktopCount; $i++) {
    Write-Log -Message "--- Start processing image" -LogFileName $logfilepath
    # Get image data from Bing json
    $imageData = $bingimagedata.images[$i]
    # Get image file name from image data
    $baseImageFileName = [System.Web.HttpUtility]::ParseQueryString(([uri]::new("http://www.bing.com$($imageData.urlbase)")).Query).Get("id")
    $imageFileName = $baseImageFileName + ".jpg"
    $imageFullFileName = Join-Path -path $wallpaperdir -ChildPath $imageFileName
    Write-Log -Message "Image file name is: $imageFileName" -LogFileName $logfilepath
    # Download image
    Invoke-WebRequest -Uri "https://www.bing.com$($imageData.url)" -OutFile $imageFullFileName
    $imageFile = Get-Item -Path $imageFullFileName
    Write-Log -Message "Downloaded image file: $($imageFile.Name)" -LogFileName $logfilepath
    # Get image description
    $imageText = $imageData | Select-Object @{l='Title';e={$_.title}},@{l='Description';e={$_.copyright}},@{l='Date';e={[System.DateTime]::ParseExact($_.enddate, "yyyyMMdd", $null)}}
    $imageHistory += [pscustomobject]@{Date = [DateTime]::now.tostring('yyyy-MMM-dd, ddd'); FileName = $baseImageFileName; Title = $imageText.Title; Description = $imageText.Description; ImageDate = $imageText.Date.tostring('yyyy-MMM-dd')}
    # Convert image
    $convertedImageFile = Convert-Image -imagefile $imageFile -imagetext $imageText
    # Set wallpaper
    Write-Log -Message "Setting wallpaper for Desktop: $i to file $($convertedImageFile.FullName)" -LogFileName $logfilepath
    $desktop = Get-Desktop $i
    Set-DesktopWallpaper -Desktop $desktop -Path $convertedImageFile
    # Delete image file
    # $convertedImageFile | Remove-Item
    $imageFile | Remove-Item
    Write-Log -Message "--- Finished processing image" -LogFileName $logfilepath
}

Write-Output $imageHistory | ConvertTo-Csv -NoHeader | Add-Content -Path $imageHistoryFile