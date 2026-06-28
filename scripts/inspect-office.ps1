$ErrorActionPreference = "Continue"
$path = 'C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\OFFICE.dll'
Write-Host "Loading: $path"
$asm = [System.Reflection.Assembly]::LoadFrom($path)
foreach ($t in $asm.GetTypes()) {
    if ($t.Name -match 'CTP|CustomTaskPane|TaskPane') {
        Write-Host $t.FullName
    }
}
Write-Host "---"
Write-Host "Done"
