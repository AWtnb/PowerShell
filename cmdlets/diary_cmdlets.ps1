
<# ==============================

cmdlets for treating diary

============================== #>

function New-DiaryTemplate {
    param (
        [int]$year,
        [int]$month
    )
    $template = @"
朝：
昼：（）
夜：（）

"@
    $weeks = @("Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat")
    if (-not $year) {
        $year = (Get-Date).Year
    }
    if (-not $month) {
        $month = (Get-Date).AddMonths(1).Month
    }
    $dir = ($pwd.ProviderPath) | Join-Path -ChildPath ("{0:d4}_{1:d2}" -f $year, $month)
    if (Test-Path $dir) {
        return
    }
    New-Item $dir -ItemType Directory
    1..31 |ForEach-Object {
        $d = Get-Date -Year $year -Month $month -Day $_
        if ($d.Month -eq $month) {
            $ts = Get-Date $d -Format "yyyy/MM/dd ddd曜"
            $fn = (Get-Date $d -Format "yyyy-MM-dd") + "_" + $weeks[$d.DayOfWeek.value__] + ".txt"
            @(
                $ts,
                $template,
                (-split "・" * 10),
                ("――" * 10)
            ) | Out-File -Path ($dir | Join-Path -ChildPath $fn) -Encoding utf8NoBOM
            "creating '{0}'" -f $fn | Write-Host
        }
    }
}


function Invoke-ConcDiary {
    <#
        .EXAMPLE
        ls | Invoke-ConcDiary
    #>
    param (
        [string]$outName
    )
    $files = $input | Where-Object {$_.Extension -eq ".txt"}
    if (-not $outName) {
        $outName = ($files | Select-Object -First 1).Basename.SubString(0, 7) -replace "[^\d]"
    }
    $files | ForEach-Object {
        "+ '{0}'" -f $_.Name | Write-Host
        Get-Content $_
    } | Out-File -Path ("Diary_{0}.txt" -f $outName) -Encoding utf8NoBOM
}