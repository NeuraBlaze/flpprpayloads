# Debug indítása
Start-Transcript -Path "$env:TEMP\flipper_local_debug.log" -Append
Write-Output "=== Levelek mentése Asztalra elkezdve ==="

# Dátumtartomány
$startDate = (Get-Date).AddDays(-7)
Write-Output "Start Date: $startDate"

# Fájlútvonal az Asztalra
$desktop = [Environment]::GetFolderPath("Desktop")
$outFile = Join-Path $desktop "Outlook_levelek_7nap.txt"

# Outlook COM API kapcsolat
try {
    $Outlook = New-Object -ComObject Outlook.Application
    $Namespace = $Outlook.GetNamespace("MAPI")
    $Inbox = $Namespace.GetDefaultFolder(6)
    $allItems = $Inbox.Items
    Write-Output "Inbox total: $($allItems.Count)"
} catch {
    Write-Output "❌ Outlook hibás vagy nem elérhető: $_"
    Stop-Transcript
    exit
}

# Szűrés
$filtered = @()
foreach ($item in $allItems) {
    try {
        if ($item.ReceivedTime -gt $startDate) {
            $filtered += $item
        }
    } catch {}
}
Write-Output "Szűrt levelek: $($filtered.Count)"

# Mentés fájlba
try {
    $filtered | ForEach-Object {
        "Subject: $($_.Subject)`nFrom: $($_.SenderName)`nDate: $($_.ReceivedTime)`nBody:`n$($_.Body)`n---`n"
    } | Set-Content -Path $outFile
    Write-Output "✅ Levelek elmentve ide: $outFile"
} catch {
    Write-Output "❌ Mentési hiba: $_"
}

# Debug lezárás
Stop-Transcript

