$css = Get-Content 'C:\Users\qinju\Desktop\AIProject\DeepExcel\src\DeepExcel.AddIn\WebViewAssets\assets\index.css' -Raw
$idx = $css.IndexOf('.send-button')
$len = $css.Length - $idx
Write-Host "Length: $($css.Length), idx: $idx, remaining: $len"
if ($idx -ge 0) {
    Write-Host $css.Substring($idx)
}
