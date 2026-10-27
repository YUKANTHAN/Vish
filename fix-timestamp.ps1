param($file = "IMG_20261023_151938.jpg.jpeg")
$shell = New-Object -ComObject Shell.Application
$folder = $shell.Namespace((Get-Item $file).DirectoryName)
$item = $folder.ParseName((Get-Item $file).Name)
$dateTaken = $folder.GetDetailsOf($item, 12)
Write-Host "Raw date string: '$dateTaken'"
$clean = $dateTaken -replace '[^\d:/APMapm ]', '' -replace '\s+', ' '
Write-Host "Cleaned: '$clean'"
$dt = $null
$formats = @("M/d/yyyy h:mm tt", "MM/dd/yyyy h:mm tt", "M/d/yyyy H:mm", "yyyy:MM:dd HH:mm:ss")
foreach ($fmt in $formats) {
    if ($clean -match '^\d') {
        try { $dt = [DateTime]::ParseExact($clean.Trim(), $fmt, $null); break } catch {}
    }
}
if (-not $dt) { $dt = Get-Date $clean }
if ($dt -and $dt.Year -gt 1900) {
    (Get-Item $file).CreationTime = $dt
    (Get-Item $file).LastWriteTime = $dt
    Write-Host "Fixed: $dt"
} else {
    Write-Host "Failed to parse date"
}