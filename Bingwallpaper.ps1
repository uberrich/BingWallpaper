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


$wallpaperdir = "$env:USERPROFILE\Pictures\BingWallpaper"
$logfile = "BingWallpaper.log"

$logfilepath = $wallpaperdir + '\' + $logfile

if (-not (Test-Path "$wallpaperdir")) {
    New-Item -ItemType Directory -Path "$wallpaperdir"
    Write-Log -Message "Created new wallpaper directory: $wallpaperdir" -LogFileName $logfilepath
}

$wpfiles = Get-ChildItem -Path $wallpaperdir

if ($wpfiles -ne $null) {
    $newestFileDate = ( $wpfiles | Sort-Object CreationTime | Select-Object -Last 1 ).CreationTime

    $daysToGet = [System.Math]::Round(([System.DateTime]::Now - $newestFileDate).TotalDays)

    if ($daysToGet -lt 1) {
        $daysToGet = 1
    }

    Write-Log -Message "Some wallpaper files already exist." -LogFileName $logfilepath
} else {
    Write-Log -Message "No wallpaper files already exist." -LogFileName $logfilepath
    $daysToGet = 5
}

Write-Log -Message "Getting $daysToGet wallpaper files." -LogFileName $logfilepath

$sFormat = [System.Drawing.StringFormat]::new()
$sFormat.alignment = [System.Drawing.StringAlignment]::Center
$sFormat.LineAlignment = [System.Drawing.StringAlignment]::Center
$font1 = [System.Drawing.Font]::new("Segoe UI",18)
$font2 = [System.Drawing.Font]::new("Segoe UI",10)
$textbrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::Black)
$fillbrush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(96,255,255,255))


$bingimagedata = Invoke-RestMethod -Uri "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=$daysToGet&mkt=en-GB" -Method Get

$bingimagedata.images | ForEach-Object {
    $imagefilename = [System.Web.HttpUtility]::ParseQueryString(([uri]::new("http://www.bing.com$($_.url)")).Query).Get("id")
    Invoke-WebRequest -Uri "https://www.bing.com$($_.url)" -OutFile "$wallpaperdir\$imagefilename"
    Write-Log -Message "Downloaded image file: $imagefilename" -LogFileName $logfilepath
    $imageText = $_ | Select-Object @{l='Title';e={$_.title}},@{l='Description';e={$_.copyright}},@{l='Date';e={[System.DateTime]::ParseExact($_.enddate, "yyyyMMdd", $null)}}
    
    $imageText | Format-List | Out-File -FilePath "$wallpaperdir\$imagefilename.txt"
    Write-Log -Message "Wrote image description file: $imagefilename.txt" -LogFileName $logfilepath

    $imageFile = Get-Item -Path "$wallpaperdir\$imagefilename"
    $bmp = [System.Drawing.Bitmap]::FromFile($imageFile)
    $image = [System.Drawing.Graphics]::FromImage($bmp)
    $SR = $bmp | Select-Object Width,Height

    $szText = $image.MeasureString(($imageText.Title + " | " + $imageText.Description), $font2)
    $rectText = [System.Drawing.RectangleF]::new((($SR.Width / 2) - ($szText.Width / 2)), ($SR.Height - ($szText.Height * 4)), $szText.Width, $szText.Height)
    $image.FillRectangle($fillbrush, $rectText)
    $image.DrawString(($imageText.Title + " | " + $imageText.Description), $font2, $textbrush, $rectText, $sFormat)
    
    $image.Dispose()
    $bmp.Save("$wallpaperdir\$imagefilename.captioned.bmp", [System.Drawing.Imaging.ImageFormat]::Bmp)
    Write-Log -Message "Wrote captioned image file: $imagefilename.captioned.bmp" -LogFileName $logfilepath
    $bmp.Dispose()
    get-item -path "$wallpaperdir\$imagefilename" | Remove-Item
    Write-Log -Message "Deleted image file: $imagefilename" -LogFileName $logfilepath
}

$wpfiles = Get-ChildItem -Path $wallpaperdir -Exclude @("*.json","*.log")

while ($wpfiles.count -gt 10) {
    Write-Log -Message "More than ten files exist. Cleaning up." -LogFileName $logfilepath
    $wpfiles | Sort-Object CreationTime | Select-Object -First 2 | Remove-Item
    Write-Log -Message "Deleted file(s): $(($wpfiles | Sort-Object CreationTime | Select-Object -First 2).Name)" -LogFileName $logfilepath
    $wpfiles = Get-ChildItem -Path $wallpaperdir -Exclude @("*.json","*.log")
}
