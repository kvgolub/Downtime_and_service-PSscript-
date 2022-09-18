. D:\ScriptWork\Downtime_and_service\Script_PowerShell\v1_1alpha\Class.ps1

function FuncOpen1($ExcelObj, [string]$d_reverse, $date) {
    $Path = [PathModule]::new($date['month'], $date['year'])

    Write-Host "Введите дату файла предыдущего отчета в формате ДД:"
    $date_report = Read-Host

    $global:ExcelWorkBook_report = $ExcelObj.Workbooks.Open($Path.path_directory + "Отчет по простоям и сервису_" + $date_report + "." + $date['month'] + "." + $date['year'] + ".xlsm")

    $ExcelWorkSheet = $ExcelWorkBook_report.Sheets.Item("Источник Рейтинг")
    $ExcelWorkSheet.Visible = $true
    $ExcelObj.Run("Выкладки")

    [int]$amount_line_old_start = $ExcelWorkBook_report.Sheets("Установочные").Cells(2, 2).Text
    $amount_line_old_end = $amount_line_old_start - 14
    $amount_line_next_start = $amount_line_old_start + 1
    $amount_line_next_end = $amount_line_old_start + 15
    $ExcelWorkSheet.Range("A" + $amount_line_old_start + ":AC" + $amount_line_old_end).Copy()
    $ExcelWorkSheet.Paste($ExcelWorkSheet.Range("A" + $amount_line_next_start))
    [datetime]$date1 = $d_reverse
    $ExcelWorkSheet.Range("B" + $amount_line_next_start + ":B" + $amount_line_next_end) = $date1
}

function FuncOpen2($ExcelObj, $d_briefly) {
    $Path = [PathModule]::new($date['month'], $date['year'])
    #$o = $Path.path_directory
    #$o2 = $Path.file()
    
    $global:source = @{}
    $index = 0
    foreach ($node in $Path.source_name_eng) {
        $source.Add($node, $ExcelObj.Workbooks.Open($Path.path_directory + $Path.path_source + $Path.source_name_rus[$index] + ".xlsx"))
        $last_sheet = $source[$node].Worksheets.Count
        $source[$node].Worksheets.Add([System.Reflection.Missing]::Value, $source[$node].Worksheets[$last_sheet]).Name = $d_briefly
        $source[$node].Worksheets("Установочные").Range("B1") = $d_briefly

        $index += 1
    }
}

function FuncOpen3($ExcelObj, $d_briefly, $d_full) {
    $Path = [PathModule]::new($date['month'], $date['year'])

    $global:ExcelWorkBook_rating = $ExcelObj.Workbooks.Open($Path.path_directory + "!!! Рейтинги (" + $Path.path_month_name + " " + $Path.path_yaer + ").xlsx")
    $last_sheet = $ExcelWorkBook_rating.Worksheets.Count
    $new_list = $ExcelWorkBook_rating.Worksheets.Add([System.Reflection.Missing]::Value, $ExcelWorkBook_rating.Worksheets[$last_sheet]).Name = $d_briefly
    $ExcelWorkBook_rating.Worksheets($new_list).Range("A1") = "Рейтинг подразделений по 4 показателям за " + $d_full
}

function FuncCopy($ExcelObj, $d_reverse, $d_briefly) {
    $Path = [PathModule]::new($date['month'], $date['year'])

    for ($index = 0; $index -lt $Path.source_name_rus.Count; $index++) {
        [string]$range = $source[$Path.source_name_eng[$index]].Sheets("Установочные").Cells(2, 2).Text
        $source[$Path.source_name_eng[$index]].Worksheets("Установочные").Range($range).Copy()

        $sheet_installation = $ExcelWorkBook_report.Worksheets("Установочные")
        [int]$amount_line_old = $sheet_installation.Cells.Item($index + 3, 2).Text
        # $workbook.Activate()
        $worksheet = $ExcelWorkBook_report.Worksheets($Path.source_name_rus[$index])
        $worksheet.Range("C" + ($amount_line_old + 1)).PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)

        [int]$amount_line_new = $sheet_installation.Cells.Item($index + 3, 2).Text
        $worksheet.Range("B" + ($amount_line_old + 1) + ":B" + $amount_line_new) = [datetime]$d_reverse
        $worksheet.Range("A" + $amount_line_old).Copy()
        $worksheet.Range("A" + ($amount_line_old + 1) + ":A" + $amount_line_new).PasteSpecial()
    }

    $sheet_rating = $ExcelWorkBook_report.Worksheets("Рейтинг")
    $sheet_rating.Activate()
    $sheet_rating.Range("R2") = [datetime]$d_reverse
    $ExcelObj.Run("Сортировка_рейтинга")
    $sheet_rating.Range("A23:B36").Copy()

    $ExcelWorkBook_rating.Worksheets($d_briefly).Range("A3").PasteSpecial([Microsoft.Office.Interop.Excel.XlPasteType]::xlPasteValues)
    $sort_rating = $ExcelWorkBook_rating.Worksheets($d_briefly)
    $sort_rating.Range("A3:B16").Sort($sort_rating.Range("A3:A16"))
}

function FuncClose($ExcelObj, $d_full) {
    $Path = [PathModule]::new($date['month'], $date['year'])

    $ExcelWorkSheet = $ExcelWorkBook_report.Sheets.Item("Источник Рейтинг")
    $ExcelWorkSheet.Activate()
    $ExcelObj.Run("Выкладки")
    $ExcelWorkSheet.Visible = $false
    # Format-List -Property Name, Index -InputObject $ExcelWorkBook_report.Sheets.Item("Рейтинг")
    $ExcelWorkBook_report.Sheets[1].Activate()
    $ExcelObj.DisplayAlerts = $false
    $ExcelWorkBook_report.SaveAs($Path.path_directory + "Отчет по простоям и сервису_" + $d_full + ".xlsm", [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled)
    $ExcelWorkBook_report.Close()

    # foreach ($key in $source.Keys) {$source[$key].Save() $source[$key].Close()}

    foreach ($node in $Path.source_name_eng) {
        $source[$node].Save()
        $source[$node].Close()
    }

    $ExcelWorkBook_rating.Save()
    $ExcelWorkBook_rating.Close()
}