$wallpaperdir = "$env:USERPROFILE\Pictures\BingWallpaper"

if (-not (Test-Path "$wallpaperdir")) {
    New-Item -ItemType Directory -Path "$wallpaperdir"
}

$wpfiles = Get-ChildItem -Path $wallpaperdir

if ($wpfiles -ne $null) {
    $newestFileDate = ( $wpfiles | Sort-Object CreationTime | Select-Object -Last 1 ).CreationTime

    $daysToGet = [System.Math]::Round(([System.DateTime]::Now - $newestFileDate).TotalDays)

    if ($daysToGet -lt 1) {
        $daysToGet = 1
    }
} else {
    $daysToGet = 5
}

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
    $imageText = $_ | Select-Object @{l='Title';e={$_.title}},@{l='Description';e={$_.copyright}},@{l='Date';e={[System.DateTime]::ParseExact($_.enddate, "yyyyMMdd", $null)}}
    
    $imageText | Format-List | Out-File -FilePath "$wallpaperdir\$imagefilename.txt"

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
    $bmp.Dispose()
    get-item -path "$wallpaperdir\$imagefilename" | Remove-Item
}

$wpfiles = Get-ChildItem -Path $wallpaperdir -Exclude "*.json"

while ($wpfiles.count -gt 10) {
    $wpfiles | Sort-Object CreationTime | Select-Object -First 2 | Remove-Item
    $wpfiles = Get-ChildItem -Path $wallpaperdir
}
