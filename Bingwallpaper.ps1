$bingimagedata = Invoke-RestMethod -Uri "https://www.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1&mkt=en-GB" -Method Get

if (-not (Test-Path "$env:USERPROFILE\Pictures\BingWallpaper")) {
    New-Item -ItemType Directory -Path "$env:USERPROFILE\Pictures\BingWallpaper"
}

Set-Location -Path "$env:USERPROFILE\Pictures\BingWallpaper"

$imagefilename = [System.Web.HttpUtility]::ParseQueryString(([uri]::new("http://www.bing.com$($bingimagedata.images[0].url)")).Query).Get("id")

Invoke-WebRequest -Uri "https://www.bing.com$($bingimagedata.images[0].url)" -OutFile $imagefilename
$bingimagedata.images[0] | Format-List title,copyright | out-file -FilePath "$imagefilename.txt"
