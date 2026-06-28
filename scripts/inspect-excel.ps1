$ErrorActionPreference = "Continue"
$path = 'C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\Microsoft.Office.Interop.Excel.dll'
Write-Host "Loading: $path"
$asm = [System.Reflection.Assembly]::LoadFrom($path)
foreach ($t in $asm.GetTypes()) {
    if ($t.Name -eq 'Application' -or $t.Name -eq '_Application') {
        Write-Host "=== $($t.FullName) ==="
        $props = $t.GetProperties([System.Reflection.BindingFlags]'Public,NonPublic,Instance,DeclaredOnly')
        foreach ($p in $props) {
            if ($p.Name -match 'CTP|TaskPane|Custom') {
                Write-Host "  PROP: $($p.Name) : $($p.PropertyType)"
            }
        }
        $methods = $t.GetMethods([System.Reflection.BindingFlags]'Public,NonPublic,Instance,DeclaredOnly')
        foreach ($m in $methods) {
            if ($m.Name -match 'CTP|TaskPane|Custom') {
                Write-Host "  METHOD: $($m.Name)"
            }
        }
    }
}
Write-Host "---"
Write-Host "Done"
