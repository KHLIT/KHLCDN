# GenerateCSV.ps1

$RepoBase = "https://media.githubusercontent.com/media/KHLIT/KHLCDN/refs/heads/main/Video"
$SourcePath = "C:\TV Media"
$OutputFile = "C:\TV Media\Karndean_Exhibit_Playlist.csv"

# Create Shell object once
$shell = New-Object -ComObject Shell.Application

# Add CSV header
$Lines = @(
    "Name,URL,Duration,UpdateInte,StartOn,EndBy,Cache"
)

$Exhibit = 1

Get-ChildItem -Path $SourcePath -Filter "*.mp4" -File |
    Sort-Object Name |
    ForEach-Object {

        $folder = $shell.Namespace($_.DirectoryName)
        $file = $folder.ParseName($_.Name)

        # Get video duration from Explorer metadata
        $videoLength = $folder.GetDetailsOf($file, 27)

        # URL encode filename
        $urlName = [System.Uri\]::EscapeDataString($_.Name)

        $Lines += (
            "Exhibit{0:D2},{1}/{2},{3},01:00:00,01/01/2018 00:00,12/31/2030 23:59,yes" -f `
            $Exhibit,
            $RepoBase,
            $urlName,
            $videoLength
        )

        $Exhibit++
    }

$Lines | Set-Content -Path $OutputFile -Encoding UTF8

Write-Host "CSV created at $OutputFile" -ForegroundColor Green
Write-Host "Entries: $($Exhibit - 1)"