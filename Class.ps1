class PathModule {
    [string]$private:path_month
    [string]$private:path_yaer
    [hashtable]$private:month_name = @{
        '01' = "Январь";
        '02' = "Февраль";
        '03' = "Март";
        '04' = "Апрель";
        '05' = "Май";
        '06' = "Июнь";
        '07' = "Июль";
        '08' = "Август";
        '09' = "Сентярь";
        '10' = "Октябрь";
        '11' = "Ноябрь";
        '12' = "Декабрь"
    }
    [string]$private:path_month_name
    [string]$path_directory
    [string]$path_source

    PathModule($p_m, $p_y) {
        $this.path_month = $p_m
        $this.path_month_name = $this.month_name[$this.path_month]
        $this.path_yaer = $p_y
        $this.path_directory = "D:\ScriptWork\Работа\2. Отчеты\1. Ежедневный\4. Простои и сервис\" + $this.path_yaer + "\" + $this.path_month + ". " + $this.path_month_name + "\"
        $this.path_source = "Исходники из 1С (" + $this.path_month_name + " " + $this.path_yaer + ")\"
    }

    [array]$source_name_eng = @(
        'ExcelWorkBook_operators_CA'
        'ExcelWorkBook_operators_CC'
        'ExcelWorkBook_technician'
        'ExcelWorkBook_revenue'
        'ExcelWorkBook_amount'
        'ExcelWorkBook_not_connection'
        'ExcelWorkBook_not_work'
    )

    [array]$source_name_rus = 
        "1.1. Операторы_(1С) (Все ТА)",
        "1.2. Операторы_(1С) (КУ)",
        "2. Техники_(1С)",
        "3. %_потер_выручки_(1С)",
        "4. Кол-во_ТА_(1С)",
        "5. ТА_без_связи_(Вин)",
        "6. Простои_(Вин)"

    [hashtable]file() {
        $source = @{}
        $index = 0
        foreach($node in $this.source_name_rus) {
            $source.Add($this.source_name_eng[$index], $node + ".xlsx")
            $index += 1
        }
        return($source)
    }
}