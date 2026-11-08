# Update-Data.ps1
# Reads the first sheet of data\sales_data.xlsx (through Excel, read-only) and writes the new data
# into index.html (the data block only - nothing else in the page is touched).
# Normal use: double-click Update-Data.bat in the project folder.
# Works with Windows PowerShell 5.1 and PowerShell 7.
param([string]$Excel = '')

$ErrorActionPreference = 'Stop'

if ($PSScriptRoot) { $Here = $PSScriptRoot } else { $Here = (Get-Location).Path }
$Root = Split-Path -Parent $Here
if ($Excel -eq '') { $Excel = Join-Path $Root 'data\sales_data.xlsx' }
$Page = Join-Path $Root 'index.html'
$Fields = @('Order ID','Order Date','Region','Branch','Category','Product','Channel','Segment','Units','Revenue','Cost')
$Inv = [System.Globalization.CultureInfo]::InvariantCulture

function Get-Cell($v, $r, $c) {
    if ($c -lt 1) { return $null }
    return $v.GetValue($r, $c)
}

function Get-Text($v, $r, $c) {
    $x = Get-Cell $v $r $c
    if ($x -eq $null) { return '' }
    return ([string]$x).Trim()
}

function Get-Number($v, $r, $c) {
    $x = Get-Cell $v $r $c
    if ($x -eq $null) { return 0.0 }
    if ($x -is [double]) { return [double]$x }
    $d = 0.0
    if ([double]::TryParse(([string]$x), [System.Globalization.NumberStyles]::Any, $Inv, [ref]$d)) { return $d }
    return 0.0
}

function Get-DateText($v, $r, $c) {
    $x = Get-Cell $v $r $c
    if ($x -eq $null) { return '' }
    if ($x -is [double]) { return [DateTime]::FromOADate($x).ToString('yyyy-MM-dd', $Inv) }
    $t = [DateTime]::MinValue
    if ([DateTime]::TryParse(([string]$x), $Inv, [System.Globalization.DateTimeStyles]::None, [ref]$t)) { return $t.ToString('yyyy-MM-dd', $Inv) }
    return ''
}

function Esc-Json([string]$s) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $s.ToCharArray()) {
        $code = [int]$ch
        if ($ch -eq '"') { [void]$sb.Append('\"') }
        elseif ($ch -eq '\') { [void]$sb.Append('\\') }
        elseif ($ch -eq '/') { [void]$sb.Append('\/') }
        elseif ($code -lt 32) { [void]$sb.Append(('\u{0:x4}' -f $code)) }
        else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

# Turns the Excel cell grid (1-based; row 1 = headers) into the JSON rows used by the page.
function Convert-Sheet($v) {
    $n = $v.GetLength(0)
    $m = $v.GetLength(1)
    $idx = @{}
    foreach ($f in $Fields) { $idx[$f] = 0 }
    for ($c = 1; $c -le $m; $c++) {
        $h = Get-Text $v 1 $c
        foreach ($f in $Fields) {
            if ($h.ToLower() -eq $f.ToLower()) { $idx[$f] = $c }
        }
    }
    foreach ($f in $Fields) {
        if ($idx[$f] -eq 0) { throw ("Column '" + $f + "' was not found in the first sheet. Keep the headers in row 1.") }
    }
    $sb = New-Object System.Text.StringBuilder
    $count = 0
    for ($i = 2; $i -le $n; $i++) {
        $id = Get-Text $v $i $idx['Order ID']
        if ($id -eq '') { continue }
        if ($count -gt 0) { [void]$sb.Append(',') }
        [void]$sb.Append('[')
        [void]$sb.Append('"' + (Esc-Json $id) + '",')
        [void]$sb.Append('"' + (Get-DateText $v $i $idx['Order Date']) + '",')
        [void]$sb.Append('"' + (Esc-Json (Get-Text $v $i $idx['Region'])) + '",')
        [void]$sb.Append('"' + (Esc-Json (Get-Text $v $i $idx['Branch'])) + '",')
        [void]$sb.Append('"' + (Esc-Json (Get-Text $v $i $idx['Category'])) + '",')
        [void]$sb.Append('"' + (Esc-Json (Get-Text $v $i $idx['Product'])) + '",')
        [void]$sb.Append('"' + (Esc-Json (Get-Text $v $i $idx['Channel'])) + '",')
        [void]$sb.Append('"' + (Esc-Json (Get-Text $v $i $idx['Segment'])) + '",')
        [void]$sb.Append([Math]::Round((Get-Number $v $i $idx['Units'])).ToString('0', $Inv) + ',')
        [void]$sb.Append([Math]::Round((Get-Number $v $i $idx['Revenue']), 2).ToString('0.##', $Inv) + ',')
        [void]$sb.Append([Math]::Round((Get-Number $v $i $idx['Cost']), 2).ToString('0.##', $Inv))
        [void]$sb.Append(']')
        $count++
    }
    return @{ Json = $sb.ToString(); Count = $count }
}

function Read-Excel {
    if (-not (Test-Path $Excel)) { throw "Excel file not found: $Excel" }
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    try {
        $wb   = $xl.Workbooks.Open($Excel, 0, $true)          # read-only: your file is never locked or changed
        $ws   = $wb.Worksheets.Item(1)
        $last = $ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row
        $lcol = $ws.Cells.Item(1, $ws.Columns.Count).End(-4159).Column
        $v    = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($last, $lcol)).Value2
        $res  = Convert-Sheet $v
        $wb.Close($false)
    }
    finally {
        $xl.Quit()
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
    }
    return $res
}

function Write-Page($res) {
    if (-not (Test-Path $Page)) { throw "index.html not found: $Page" }
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm'
    $data = '{"built":"' + $stamp + '","rows":[' + $res.Json + ']}'
    $t = [System.IO.File]::ReadAllText($Page, [System.Text.Encoding]::UTF8)
    $open = '<script id="rawdata" type="application/json">'
    $a = $t.IndexOf($open)
    if ($a -lt 0) { throw 'index.html does not contain the data block.' }
    $a = $a + $open.Length
    $b = $t.IndexOf('</script>', $a)
    $t = $t.Substring(0, $a) + $data + $t.Substring($b)
    [System.IO.File]::WriteAllText($Page, $t, (New-Object System.Text.UTF8Encoding($false)))
    Write-Host ("{0}  updated: {1} orders" -f (Get-Date -Format 'HH:mm:ss'), $res.Count)
}

function Run {
    try {
        Write-Host ('Reading ' + (Split-Path -Leaf $Excel) + ' ...')
        $res = Read-Excel
        Write-Page $res
    }
    catch {
        Write-Host ''
        Write-Host ('ERROR: ' + $_.Exception.Message) -ForegroundColor Red
        exit 1
    }
}

if ($MyInvocation.InvocationName -ne '.') { Run }
