# Fun project that scrapes xkcd.com and downloads all the comics.

# main
New-Item -Path "$PSScriptRoot\XKCD" -ItemType "Directory" -ErrorAction "SilentlyContinue"

$baseUrl = 'https://xkcd.com'
$url = $baseUrl
while ($url -notlike "*#")
{
    Write-Host "Downloading page: $url" -ForegroundColor "DarkCyan"

    try
    {
        $response = Invoke-WebRequest -Uri $url -ErrorVariable "responseError"
    }
    catch
    {
        Write-Host "$($responseError[0].Message)" -ForegroundColor "Red"
        $responseCode = [int]$_.Exception.Response.StatusCode
        Write-Host "Response code is: $responseCode" -ForegroundColor "Red"
        exit
    }

    $html = ConvertFrom-Html -Content $response
    $comicElem = $html.SelectSingleNode('/html[1]/body[1]/div[2]/div[2]/img[1]')

    if ($null -eq $comicElem)
    {
        Write-Warning "Could not find comic image."    
    }
    else
    {
        $comicUrl = "https:$($comicElem.Attributes[0].Value)"
        $uri = New-Object System.Uri($comicUrl)
        $endPart = $uri.Segments[-1]

        try
        {
            # Downloads image into path specified in OutFile.
            Invoke-WebRequest -Uri $comicUrl -OutFile "$PSScriptRoot\XKCD\$endPart" -ErrorVariable "responseError"
        }
        catch
        {
            Write-Host "$($responseError[0].Message)" -ForegroundColor "Red"
            $responseCode = [int]$_.Exception.Response.StatusCode
            Write-Host "Response code is: $responseCode" -ForegroundColor "Red"
            exit
        }     
    }

    $prevLink = $html.SelectSingleNode('/html[1]/body[1]/div[2]/ul[1]/li[2]/a[1]')
    $matchInfo = $prevLink.OuterHtml | Select-String -Pattern '(?<=href=").*?(?=")'
    $href = $matchInfo.Matches[0].Value
    $url = $baseUrl + $href
}

Write-Host "Done" -ForegroundColor "Green"