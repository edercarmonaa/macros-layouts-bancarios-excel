param(
    [string]$WorkbookPattern = "*.xlsm",
    [string]$OutputDirectory = "vba"
)

$ErrorActionPreference = "Stop"

function Get-SafeName {
    param([string]$Name)

    $safe = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $safe = $safe -replace '[^\w.-]+', '-'
    $safe = $safe.Trim('-')
    return $safe
}

function Get-ExtensionForComponent {
    param([int]$Type)

    switch ($Type) {
        1 { ".bas" }  # Standard module
        2 { ".cls" }  # Class module
        3 { ".frm" }  # UserForm
        100 { ".cls" } # Document module
        default { ".txt" }
    }
}

$root = (Resolve-Path ".").Path
$outputRoot = Join-Path $root $OutputDirectory
New-Item -ItemType Directory -Force -Path $outputRoot | Out-Null

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    $workbooks = Get-ChildItem -Path $root -Filter $WorkbookPattern -File

    foreach ($file in $workbooks) {
        $safeWorkbookName = Get-SafeName $file.Name

        Write-Host "Exporting VBA from $($file.Name)"
        $workbook = $excel.Workbooks.Open($file.FullName, $false, $true)

        try {
            $components = @($workbook.VBProject.VBComponents)

            if ($components.Count -eq 0) {
                Write-Warning "  No exportable VBA components were exposed by Excel."
                continue
            }

            $targetDir = Join-Path $outputRoot $safeWorkbookName
            New-Item -ItemType Directory -Force -Path $targetDir | Out-Null

            foreach ($component in $components) {
                $extension = Get-ExtensionForComponent ([int]$component.Type)
                $targetFile = Join-Path $targetDir ($component.Name + $extension)
                $component.Export($targetFile)
                Write-Host "  -> $targetFile"
            }
        }
        finally {
            $workbook.Close($false)
        }
    }
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
}

Write-Host "VBA export completed in $outputRoot"
