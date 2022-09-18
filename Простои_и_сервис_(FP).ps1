. D:\ScriptWork\Downtime_and_service\Script_PowerShell\v1_1alpha\Function.ps1

$local:ExcelObj = New-Object -ComObject Excel.Application
$ExcelObj.Visible = $true
$ExcelObj.WindowState = 'xlMaximized'

Write-Host "Введите дату отчета в формате ГГГГ.ММ.ДД:"
$local:d_reverse = Read-Host

$local:date = @{
    'day' = $d_reverse.Substring(8, 2)
    'month' = $d_reverse.Substring(5, 2)
    'year' =$d_reverse.Substring(0, 4)
}

$local:d_full = $date['day'] + "." + $date['month'] + "." + $date['year']
$local:d_briefly = $date['day'] + "." + $date['month']

while ($true) {
    Write-Host "Выберите функцию:"
    Write-Host "1. Открыть файлы отчета"
    Write-Host "2. Копировать сведения"
    Write-Host "3. Сохраниить и закрыть все файлы"
    Write-Host "4. Завершить скрипт"

    $v = Read-Host

    if ($v -eq "1") {
        FuncOpen1 $ExcelObj $d_reverse $date
        FuncOpen2 $ExcelObj $d_briefly
        FuncOpen3 $ExcelObj $d_briefly $d_full
    }
    elseif ($v -eq "2") {
        FuncCopy $ExcelObj $d_reverse $d_briefly
    }
    elseif ($v -eq "3") {
        FuncClose $ExcelObj $d_full
    }
    elseif ($v -eq "4") {
        $a = Get-Process -Name EXCEL
        $a | Stop-Process
        break
    }
}