$bingimagedata = Invoke-RestMethod -Uri "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-GB" -Method Get

$wallpaperdir = "$env:USERPROFILE\Pictures\BingWallpaper"

if (-not (Test-Path "$wallpaperdir")) {
    New-Item -ItemType Directory -Path "$wallpaperdir"
}

$imagefilename = [System.Web.HttpUtility]::ParseQueryString(([uri]::new("http://www.bing.com$($bingimagedata.images[0].url)")).Query).Get("id")

Invoke-WebRequest -Uri "https://www.bing.com$($bingimagedata.images[0].url)" -OutFile "$wallpaperdir\$imagefilename"
$bingimagedata.images[0] | Format-List title,copyright | out-file -FilePath "$wallpaperdir\$imagefilename.txt"

$wpfiles = Get-ChildItem -Path $wallpaperdir

$wpfiles | ForEach-Object {
    if ($_.CreationTimeUtc.CompareTo([DateTime]::UtcNow.AddDays(-3)) -eq -1 ) {
        $_ | Remove-Item
    }
}
