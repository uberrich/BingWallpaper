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

$sFormat = [System.Drawing.StringFormat]::new
$sFormat.alignment = [System.Drawing.StringAlignment]::Far

$bingimagedata = Invoke-RestMethod -Uri "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=$daysToGet&mkt=en-GB" -Method Get

$bingimagedata.images | ForEach-Object {
    $imagefilename = [System.Web.HttpUtility]::ParseQueryString(([uri]::new("http://www.bing.com$($_.url)")).Query).Get("id")
    Invoke-WebRequest -Uri "https://www.bing.com$($_.url)" -OutFile "$wallpaperdir\$imagefilename"
    $imageText = $_ | Select-Object @{l='Title';e={$_.title}},@{l='Description';e={$_.copyright}}
    
    $imageText | Format-List | Out-File -FilePath "$wallpaperdir\$imagefilename.txt"

    $imageFile = Get-Item -Path "$wallpaperdir\$imagefilename"
    $bmp = [System.Drawing.Bitmap]::FromFile($imageFile)
    $image = [System.Drawing.Graphics]::FromImage($bmp)
    $SR = $bmp | Select-Object Width,Height
}

$wpfiles = Get-ChildItem -Path $wallpaperdir

while ($wpfiles.count -gt 10) {
    $wpfiles | Sort-Object CreationTime | Select-Object -First 2 | Remove-Item
    $wpfiles = Get-ChildItem -Path $wallpaperdir
}
