@echo off
REM DeepExcel Build Script - uses csc.exe directly
setlocal

set BASEDIR=%~dp0..
set ADDINDIR=%BASEDIR%\src\DeepExcel.AddIn
set OUTDIR=%ADDINDIR%\bin\Release
set PACKAGES=%BASEDIR%\packages
set CSC=%PACKAGES%\Microsoft.Net.Compilers.3.8.0\tools\csc.exe

if not exist "%OUTDIR%" mkdir "%OUTDIR%"

echo === DeepExcel Build ===
echo.

REM Copy all dependency DLLs to output first
echo [1/3] Copying dependencies...
copy /Y "%PACKAGES%\Microsoft.Web.WebView2.1.0.2420.47\lib\net45\Microsoft.Web.WebView2.WinForms.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\Microsoft.Web.WebView2.1.0.2420.47\lib\net45\Microsoft.Web.WebView2.Core.dll" "%OUTDIR%\" >nul
if not exist "%OUTDIR%\runtimes\win-x86\native" mkdir "%OUTDIR%\runtimes\win-x86\native" >nul
if not exist "%OUTDIR%\runtimes\win-x64\native" mkdir "%OUTDIR%\runtimes\win-x64\native" >nul
copy /Y "%PACKAGES%\Microsoft.Web.WebView2.1.0.2420.47\runtimes\win-x86\native\WebView2Loader.dll" "%OUTDIR%\runtimes\win-x86\native\" >nul
copy /Y "%PACKAGES%\Microsoft.Web.WebView2.1.0.2420.47\runtimes\win-x64\native\WebView2Loader.dll" "%OUTDIR%\runtimes\win-x64\native\" >nul
copy /Y "%PACKAGES%\Microsoft.Web.WebView2.1.0.2420.47\runtimes\win-x86\native\WebView2Loader.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\System.Text.Json.8.0.0\lib\net462\System.Text.Json.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\System.Text.Encodings.Web.8.0.0\lib\net462\System.Text.Encodings.Web.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\Microsoft.Bcl.AsyncInterfaces.8.0.0\lib\net462\Microsoft.Bcl.AsyncInterfaces.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\System.Buffers.4.5.1\lib\net461\System.Buffers.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\System.Memory.4.5.5\lib\net461\System.Memory.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\System.Numerics.Vectors.4.5.0\lib\net46\System.Numerics.Vectors.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\System.Runtime.CompilerServices.Unsafe.6.0.0\lib\net461\System.Runtime.CompilerServices.Unsafe.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\System.Threading.Tasks.Extensions.4.5.4\lib\portable-net45+win8+wp8+wpa81\System.Threading.Tasks.Extensions.dll" "%OUTDIR%\" >nul
copy /Y "%PACKAGES%\System.ValueTuple.4.5.0\lib\net47\System.ValueTuple.dll" "%OUTDIR%\" >nul

REM Copy WebViewAssets
echo [2/3] Copying WebView assets...
if exist "%OUTDIR%\WebViewAssets" rmdir /S /Q "%OUTDIR%\WebViewAssets"
xcopy /E /I /Q "%ADDINDIR%\WebViewAssets" "%OUTDIR%\WebViewAssets" >nul
copy /Y "%ADDINDIR%\App.config" "%OUTDIR%\DeepExcel.AddIn.dll.config" >nul

REM Copy sidecar scripts
echo [2.5/3] Copying sidecar scripts...
if not exist "%OUTDIR%\sidecar" mkdir "%OUTDIR%\sidecar"
xcopy /E /I /Q /Y "%BASEDIR%\src\DeepExcel.Sidecar\*.py" "%OUTDIR%\sidecar" >nul

REM Compile
echo [3/3] Compiling DeepExcel.AddIn.dll...
"%CSC%" ^
  /target:library ^
  /out:"%OUTDIR%\DeepExcel.AddIn.dll" ^
  /langversion:9.0 ^
  /optimize ^
  /debug:pdbonly ^
  /define:TRACE ^
  /platform:anycpu ^
  /unsafe- ^
  /nowin32manifest ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\mscorlib.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\System.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\System.Core.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\System.Drawing.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\System.Net.Http.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\System.Windows.Forms.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\System.Xml.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\System.Xml.Linq.dll" ^
  /reference:"%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\Microsoft.CSharp.dll" ^
  /reference:"C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\Microsoft.Office.Interop.Excel.dll" ^
  /reference:"C:\Program Files\Microsoft Office\root\Office16\ADDINS\PowerPivot Excel Add-in\OFFICE.dll" ^
  /reference:"C:\Program Files (x86)\Common Files\Microsoft Shared\MSEnv\PublicAssemblies\Extensibility.dll" ^
  /reference:"%WINDIR%\assembly\GAC_MSIL\Microsoft.Vbe.Interop\15.0.0.0__71e9bce111e9429c\Microsoft.Vbe.Interop.dll" ^
  /reference:"%OUTDIR%\Microsoft.Web.WebView2.WinForms.dll" ^
  /reference:"%OUTDIR%\Microsoft.Web.WebView2.Core.dll" ^
  /reference:"%OUTDIR%\System.Text.Json.dll" ^
  /reference:"%OUTDIR%\System.Text.Encodings.Web.dll" ^
  /reference:"%OUTDIR%\Microsoft.Bcl.AsyncInterfaces.dll" ^
  /reference:"%OUTDIR%\System.Buffers.dll" ^
  /reference:"%OUTDIR%\System.Memory.dll" ^
  /reference:"%OUTDIR%\System.Numerics.Vectors.dll" ^
  /reference:"%OUTDIR%\System.Runtime.CompilerServices.Unsafe.dll" ^
  /reference:"%OUTDIR%\System.Threading.Tasks.Extensions.dll" ^
  /reference:"%OUTDIR%\System.ValueTuple.dll" ^
  /resource:"%ADDINDIR%\Resources\DeepExcelRibbon.xml",DeepExcel.AddIn.Resources.DeepExcelRibbon.xml ^
  "%ADDINDIR%\Advanced\ChartSpecEngine.cs" ^
  "%ADDINDIR%\Bridge\IExcelActions.cs" ^
  "%ADDINDIR%\Bridge\MessageBridge.cs" ^
  "%ADDINDIR%\Bridge\Messages.cs" ^
  "%ADDINDIR%\Bridge\WorkbookSession.cs" ^
  "%ADDINDIR%\Collaboration\OperationHistory.cs" ^
  "%ADDINDIR%\Collaboration\ConversationHistory.cs" ^
  "%ADDINDIR%\Config\AppConfig.cs" ^
  "%ADDINDIR%\Diagnostics\Logger.cs" ^
  "%ADDINDIR%\Executor\ExecutionResult.cs" ^
  "%ADDINDIR%\Executor\PythonExecutor.cs" ^
  "%ADDINDIR%\Executor\SnapshotManager.cs" ^
  "%ADDINDIR%\Executor\VBAExecutor.cs" ^
  "%ADDINDIR%\Performance\PerformanceOptimizer.cs" ^
  "%ADDINDIR%\Perception\RangeAnalyzer.cs" ^
  "%ADDINDIR%\Perception\RangeInfo.cs" ^
  "%ADDINDIR%\Perception\WorkbookAnalyzer.cs" ^
  "%ADDINDIR%\Perception\WorkbookStructure.cs" ^
  "%ADDINDIR%\Properties\AssemblyInfo.cs" ^
  "%ADDINDIR%\Security\CodeSandbox.cs" ^
  "%ADDINDIR%\Security\SecurityGateway.cs" ^
  "%ADDINDIR%\Security\SecurityManager.cs" ^
  "%ADDINDIR%\Sidecar\PythonSidecar.cs" ^
  "%ADDINDIR%\Sidecar\SidecarProtocol.cs" ^
  "%ADDINDIR%\Sidecar\ToolDispatcher.cs" ^
  "%ADDINDIR%\Sidecar\JsonConverters.cs" ^
  "%ADDINDIR%\TaskPaneControl.cs" ^
  "%ADDINDIR%\IRibbonCallbacks.cs" ^
  "%ADDINDIR%\ThisAddIn.cs" ^
  "%ADDINDIR%\Tools\ChartTool.cs" ^
  "%ADDINDIR%\Tools\DataCleaner.cs" ^
  "%ADDINDIR%\Tools\FormulaTool.cs" ^
  "%ADDINDIR%\Extensions\DictionaryExtensions.cs"

if %ERRORLEVEL% EQU 0 (
  echo.
  echo === Build SUCCESSFUL ===
  echo Output: %OUTDIR%\DeepExcel.AddIn.dll
  echo.
) else (
  echo.
  echo === Build FAILED ===
  exit /b 1
)

echo.
echo === ?? Python sidecar ===
powershell -ExecutionPolicy Bypass -File "%~dp0package-python.ps1"
if errorlevel 1 (
    echo [ERROR] Python ????
    exit /b 1
)

endlocal
