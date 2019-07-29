$wallpaperdir = "$env:USERPROFILE\Pictures\BingWallpaperTest"

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
$sFormat.alignment = [System.Drawing.StringAlignment]::Far
$font1 = [System.Drawing.Font]::new("Segoe UI",28)
$font2 = [System.Drawing.Font]::new("Segoe UI",14)
$brush = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(192,192,192))


$bingimagedata = Invoke-RestMethod -Uri "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=$daysToGet&mkt=en-GB" -Method Get

$bingimagedata | ConvertTo-Json | Out-File -FilePath "$wallpaperdir\bingimagedata.json"

$bingimagedata.images | ForEach-Object {
    $imagefilename = [System.Web.HttpUtility]::ParseQueryString(([uri]::new("http://www.bing.com$($_.url)")).Query).Get("id")
    Invoke-WebRequest -Uri "https://www.bing.com$($_.url)" -OutFile "$wallpaperdir\$imagefilename"
    #Start-Sleep -Seconds 10 
    $imageText = $_ | Select-Object @{l='Title';e={$_.title}},@{l='Description';e={$_.copyright}}
    
    $imageText | Format-List | Out-File -FilePath "$wallpaperdir\$imagefilename.txt"

    $imageFile = Get-Item -Path "$wallpaperdir\$imagefilename"
    $bmp = [System.Drawing.Bitmap]::FromFile($imageFile)
    $image = [System.Drawing.Graphics]::FromImage($bmp)
    $SR = $bmp | Select-Object Width,Height
    $sz = $image.MeasureString($($imageText.Title), $font1)
    $rect1 = [System.Drawing.RectangleF]::new(0,$sz.Height,($SR.Width - $sz.Height),$SR.Height)
    $rect2 = [System.Drawing.RectangleF]::new(0,($sz.Height * 2),($SR.Width - $sz.Height),$SR.Height)
    $image.DrawString($imageText.Title, $font1, $brush, $rect1, $sFormat)
    $image.DrawString($imageText.Description, $font2, $brush, $rect2, $sFormat)
    $image.Dispose()
    $bmp.Save("$wallpaperdir\$imagefilename.captioned.jpg", [System.Drawing.Imaging.ImageFormat]::Jpeg)
    $bmp.Dispose()
    get-item -path "$wallpaperdir\$imagefilename" | Remove-Item
}

$wpfiles = Get-ChildItem -Path $wallpaperdir -Exclude "*.json"

while ($wpfiles.count -gt 10) {
    $wpfiles | Sort-Object CreationTime | Select-Object -First 2 | Remove-Item
    $wpfiles = Get-ChildItem -Path $wallpaperdir
}
