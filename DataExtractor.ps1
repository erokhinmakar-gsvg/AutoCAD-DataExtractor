# AutoCAD Data Extractor
# Монолитный PowerShell скрипт для извлечения данных из DWG файлов

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ========== Словарь версий AutoCAD ==========
$versions = @{
    "2020" = "AutoCAD.Application.24"
    "2021" = "AutoCAD.Application.25"
    "2022" = "AutoCAD.Application.26"
}

# ========== Глобальные переменные ==========
$global:acadApp = $null
$global:acadDoc = $null
$global:currentData = $null       # последние результаты для экспорта
$global:currentColumns = @()      # заголовки столбцов

# ========== Вспомогательные функции ==========

# Функция повторных попыток для COM-вызовов
function Invoke-WithRetry {
    param(
        [ScriptBlock]$ScriptBlock,
        [int]$MaxRetries = 3,
        [int]$DelayMs = 500
    )
    $lastError = $null
    for ($i = 0; $i -lt $MaxRetries; $i++) {
        try {
            return & $ScriptBlock
        } catch {
            $lastError = $_
            Write-Host "Попытка $($i+1) не удалась: $($_.Exception.Message). Повтор через ${DelayMs}мс..."
            Start-Sleep -Milliseconds $DelayMs
        }
    }
    throw $lastError
}

# Подключение к AutoCAD
function Connect-AutoCAD {
    param(
        [string]$version,
        [string]$filePath
    )
    $progId = $versions[$version]
    if (-not $progId) {
        throw "Неизвестная версия AutoCAD: $version"
    }
    try {
        $acad = Invoke-WithRetry -ScriptBlock { [System.Runtime.InteropServices.Marshal]::GetActiveObject($progId) }
        Write-Host "Подключились к запущенному AutoCAD $version"
    } catch {
        try {
            $acad = New-Object -ComObject $progId
            $acad.Visible = $false
            Write-Host "Запустили новый экземпляр AutoCAD $version"
        } catch {
            throw "AutoCAD версии $version не найден на этом компьютере."
        }
    }
    try { $acad.DisplayAlerts = $false } catch { }

    try {
        $doc = Invoke-WithRetry -ScriptBlock { $acad.Documents.Open($filePath) }
        Start-Sleep -Milliseconds 1000  # даём время на загрузку
    } catch {
        throw "Не удалось открыть файл: $($_.Exception.Message)"
    }
    return @{ App = $acad; Doc = $doc }
}

function Disconnect-AutoCAD {
    param($acad, $doc)
    if ($doc) {
        try { Invoke-WithRetry -ScriptBlock { $doc.Close($false) } -MaxRetries 2 } catch { }
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($doc) | Out-Null
    }
    if ($acad) {
        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($acad) | Out-Null
    }
}

function Get-RadiusRange {
    param([double]$radius)
    for ($i = 0; $i -le 12; $i++) {
        if ($radius -ge ($i + 0.5) -and $radius -lt ($i + 1.5)) {
            return "R = $($i + 1)"
        }
    }
    return $null
}

function Sort-StringArray {
    param($array)
    [array]::Sort($array)
    return $array
}

# ========== Форма выбора слоя ==========
function Show-LayerSelectionDialog {
    param($layerList)
    $dialog = New-Object System.Windows.Forms.Form
    $dialog.Text = "Выбор слоя"
    $dialog.Size = New-Object System.Drawing.Size(300, 150)
    $dialog.StartPosition = "CenterParent"
    $dialog.FormBorderStyle = "FixedDialog"
    $dialog.MaximizeBox = $false
    $dialog.MinimizeBox = $false

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Выберите слой (или оставьте пустым для всех):"
    $label.Location = New-Object System.Drawing.Point(10, 10)
    $label.Size = New-Object System.Drawing.Size(260, 20)

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Location = New-Object System.Drawing.Point(10, 35)
    $combo.Size = New-Object System.Drawing.Size(260, 25)
    $combo.DropDownStyle = "DropDownList"
    $combo.Items.Add("") | Out-Null
    foreach ($layer in $layerList) {
        $combo.Items.Add($layer) | Out-Null
    }
    $combo.SelectedIndex = 0

    $okButton = New-Object System.Windows.Forms.Button
    $okButton.Text = "OK"
    $okButton.Location = New-Object System.Drawing.Point(100, 70)
    $okButton.Size = New-Object System.Drawing.Size(80, 25)
    $okButton.DialogResult = "OK"

    $dialog.Controls.AddRange(@($label, $combo, $okButton))
    $dialog.AcceptButton = $okButton

    if ($dialog.ShowDialog() -eq "OK") {
        return $combo.SelectedItem
    }
    return $null
}

# ========== Главная форма ==========
$mainForm = New-Object System.Windows.Forms.Form
$mainForm.Text = "AutoCAD Data Extractor"
$mainForm.Size = New-Object System.Drawing.Size(850, 600)  # уменьшенный размер
$mainForm.StartPosition = "CenterScreen"
$mainForm.BackColor = "LightGray"
$mainForm.MinimumSize = New-Object System.Drawing.Size(800, 500)

# Элементы управления
$labelVersion = New-Object System.Windows.Forms.Label
$labelVersion.Text = "Версия AutoCAD:"
$labelVersion.Location = New-Object System.Drawing.Point(10, 10)
$labelVersion.Size = New-Object System.Drawing.Size(100, 25)

$comboVersion = New-Object System.Windows.Forms.ComboBox
$comboVersion.Location = New-Object System.Drawing.Point(120, 10)
$comboVersion.Size = New-Object System.Drawing.Size(80, 25)
$comboVersion.DropDownStyle = "DropDownList"
$comboVersion.Items.AddRange(@("2020", "2021", "2022"))
$comboVersion.SelectedIndex = 0

$labelFile = New-Object System.Windows.Forms.Label
$labelFile.Text = "DWG файл:"
$labelFile.Location = New-Object System.Drawing.Point(220, 10)
$labelFile.Size = New-Object System.Drawing.Size(60, 25)

$textBoxFile = New-Object System.Windows.Forms.TextBox
$textBoxFile.Location = New-Object System.Drawing.Point(290, 10)
$textBoxFile.Size = New-Object System.Drawing.Size(350, 25)
$textBoxFile.ReadOnly = $true

$buttonBrowse = New-Object System.Windows.Forms.Button
$buttonBrowse.Text = "Обзор..."
$buttonBrowse.Location = New-Object System.Drawing.Point(650, 10)
$buttonBrowse.Size = New-Object System.Drawing.Size(80, 25)
$buttonBrowse.Add_Click({
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Filter = "DWG файлы (*.dwg)|*.dwg|Все файлы (*.*)|*.*"
    if ($openFileDialog.ShowDialog() -eq "OK") {
        $textBoxFile.Text = $openFileDialog.FileName
    }
})

$buttonLength = New-Object System.Windows.Forms.Button
$buttonLength.Text = "Длины"
$buttonLength.Location = New-Object System.Drawing.Point(10, 50)
$buttonLength.Size = New-Object System.Drawing.Size(80, 30)
$buttonLength.Add_Click({ Calculate-Length })

$buttonArea = New-Object System.Windows.Forms.Button
$buttonArea.Text = "Площади"
$buttonArea.Location = New-Object System.Drawing.Point(100, 50)
$buttonArea.Size = New-Object System.Drawing.Size(80, 30)
$buttonArea.Add_Click({ Calculate-Area })

$buttonArc = New-Object System.Windows.Forms.Button
$buttonArc.Text = "Дуги"
$buttonArc.Location = New-Object System.Drawing.Point(190, 50)
$buttonArc.Size = New-Object System.Drawing.Size(80, 30)
$buttonArc.Add_Click({ Calculate-Arc })

$buttonBlock = New-Object System.Windows.Forms.Button
$buttonBlock.Text = "Блоки"
$buttonBlock.Location = New-Object System.Drawing.Point(280, 50)
$buttonBlock.Size = New-Object System.Drawing.Size(80, 30)
$buttonBlock.Add_Click({ Calculate-Block })

$buttonSaveCsv = New-Object System.Windows.Forms.Button
$buttonSaveCsv.Text = "Сохранить CSV"
$buttonSaveCsv.Location = New-Object System.Drawing.Point(380, 50)
$buttonSaveCsv.Size = New-Object System.Drawing.Size(120, 30)
$buttonSaveCsv.Add_Click({ Export-CsvFile })

$buttonSaveExcel = New-Object System.Windows.Forms.Button
$buttonSaveExcel.Text = "Экспорт Excel"
$buttonSaveExcel.Location = New-Object System.Drawing.Point(510, 50)
$buttonSaveExcel.Size = New-Object System.Drawing.Size(120, 30)
$buttonSaveExcel.Add_Click({ Export-Excel })

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Location = New-Object System.Drawing.Point(10, 90)
$statusLabel.Size = New-Object System.Drawing.Size(500, 20)
$statusLabel.Text = "Готов"

# ========== TabControl с вкладками ==========
$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(10, 120)
$tabControl.Size = New-Object System.Drawing.Size(810, 400)
$tabControl.Anchor = "Top, Bottom, Left, Right"

# Вкладка Длины
$tabLength = New-Object System.Windows.Forms.TabPage
$tabLength.Text = "Длины"
$dataGridViewLength = New-Object System.Windows.Forms.DataGridView
$dataGridViewLength.Dock = "Fill"
$dataGridViewLength.AutoSizeColumnsMode = "Fill"
$dataGridViewLength.ReadOnly = $true
$dataGridViewLength.AllowUserToAddRows = $false
$dataGridViewLength.RowHeadersVisible = $false
$tabLength.Controls.Add($dataGridViewLength)

# Вкладка Площади
$tabArea = New-Object System.Windows.Forms.TabPage
$tabArea.Text = "Площади"
$dataGridViewArea = New-Object System.Windows.Forms.DataGridView
$dataGridViewArea.Dock = "Fill"
$dataGridViewArea.AutoSizeColumnsMode = "Fill"
$dataGridViewArea.ReadOnly = $true
$dataGridViewArea.AllowUserToAddRows = $false
$dataGridViewArea.RowHeadersVisible = $false
$tabArea.Controls.Add($dataGridViewArea)

# Вкладка Дуги
$tabArc = New-Object System.Windows.Forms.TabPage
$tabArc.Text = "Дуги"
$dataGridViewArc = New-Object System.Windows.Forms.DataGridView
$dataGridViewArc.Dock = "Fill"
$dataGridViewArc.AutoSizeColumnsMode = "Fill"
$dataGridViewArc.ReadOnly = $true
$dataGridViewArc.AllowUserToAddRows = $false
$dataGridViewArc.RowHeadersVisible = $false
$tabArc.Controls.Add($dataGridViewArc)

# Вкладка Блоки
$tabBlock = New-Object System.Windows.Forms.TabPage
$tabBlock.Text = "Блоки"
$dataGridViewBlock = New-Object System.Windows.Forms.DataGridView
$dataGridViewBlock.Dock = "Fill"
$dataGridViewBlock.AutoSizeColumnsMode = "Fill"
$dataGridViewBlock.ReadOnly = $true
$dataGridViewBlock.AllowUserToAddRows = $false
$dataGridViewBlock.RowHeadersVisible = $false
$tabBlock.Controls.Add($dataGridViewBlock)

$tabControl.TabPages.AddRange(@($tabLength, $tabArea, $tabArc, $tabBlock))

# Кнопка закрытия
$buttonClose = New-Object System.Windows.Forms.Button
$buttonClose.Text = "Закрыть"
$buttonClose.Location = New-Object System.Drawing.Point(730, 530)
$buttonClose.Size = New-Object System.Drawing.Size(100, 30)
$buttonClose.Anchor = "Bottom, Right"
$buttonClose.Add_Click({ $mainForm.Close() })

# ========== Всплывающие подсказки для кнопок ==========
$tooltip = New-Object System.Windows.Forms.ToolTip
$tooltip.SetToolTip($buttonLength,    "Вычисление суммарной длины линейных объектов (отрезков, дуг, полилиний) с группировкой по слоям.")
$tooltip.SetToolTip($buttonArea,      "Вычисление суммарной площади штриховок (Hatch) с группировкой по слоям. Предупреждает о штриховках с нулевой площадью.")
$tooltip.SetToolTip($buttonArc,       "Вычисление суммарной длины дуг, сгруппированных по слоям и диапазонам радиусов (R1–R12). Полилинии предварительно взрываются.")
$tooltip.SetToolTip($buttonBlock,     "Подсчёт количества вхождений блоков с группировкой по слоям и состояниям видимости (для динамических блоков).")
$tooltip.SetToolTip($buttonSaveCsv,   "Экспорт данных активной вкладки в CSV-файл (разделитель ';', кодировка UTF-8).")
$tooltip.SetToolTip($buttonSaveExcel, "Экспорт данных активной вкладки в Excel-файл (.xlsx) с форматированием.")
$tooltip.SetToolTip($buttonBrowse,    "Выбрать DWG-файл для анализа.")
$tooltip.SetToolTip($buttonClose,     "Закрыть программу.")

$mainForm.Controls.AddRange(@(
    $labelVersion, $comboVersion,
    $labelFile, $textBoxFile, $buttonBrowse,
    $buttonLength, $buttonArea, $buttonArc, $buttonBlock,
    $buttonSaveCsv, $buttonSaveExcel,
    $statusLabel,
    $tabControl,
    $buttonClose
))

# Обработчик смены вкладки: обновляем $global:currentData данными активной таблицы
$tabControl.Add_SelectedIndexChanged({
    $selectedTab = $tabControl.SelectedTab
    $grid = $selectedTab.Controls[0]  # предполагаем, что DataGridView первый и единственный элемент на вкладке
    $global:currentData = @()
    foreach ($row in $grid.Rows) {
        if (-not $row.IsNewRow) {
            $obj = New-Object PSObject
            for ($i = 0; $i -lt $grid.Columns.Count; $i++) {
                $colName = $grid.Columns[$i].Name
                $value = $row.Cells[$i].Value
                $obj | Add-Member -MemberType NoteProperty -Name $colName -Value $value
            }
            $global:currentData += $obj
        }
    }
})

# ========== Функции расчетов ==========

function Calculate-Length {
    $global:currentColumns = @("Длина", "Слой")
    $file = $textBoxFile.Text
    if ([string]::IsNullOrWhiteSpace($file)) {
        [System.Windows.Forms.MessageBox]::Show("Выберите DWG файл!", "Ошибка")
        return
    }
    $version = $comboVersion.SelectedItem

    # Показываем индикатор загрузки
    $statusLabel.Text = "Загрузка и извлечение данных..."
    $mainForm.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $conn = Connect-AutoCAD -version $version -filePath $file
        $acad = $conn.App
        $doc = $conn.Doc
        $model = Invoke-WithRetry -ScriptBlock { $doc.ModelSpace }

        if (-not $model) {
            throw "Не удалось получить ModelSpace"
        }

        $dict = @{}
        $count = 0

        foreach ($entity in $model) {
            try {
                $type = Invoke-WithRetry -ScriptBlock { $entity.ObjectName } -MaxRetries 2
                $layer = Invoke-WithRetry -ScriptBlock { $entity.Layer } -MaxRetries 2
                $length = $null

                if ($type -eq "AcDbLine") {
                    $sp = Invoke-WithRetry -ScriptBlock { $entity.StartPoint }
                    $ep = Invoke-WithRetry -ScriptBlock { $entity.EndPoint }
                    $dx = $ep[0] - $sp[0]
                    $dy = $ep[1] - $sp[1]
                    $dz = $ep[2] - $sp[2]
                    $length = [math]::Sqrt($dx*$dx + $dy*$dy + $dz*$dz)
                }
                elseif ($type -eq "AcDbArc") {
                    $length = Invoke-WithRetry -ScriptBlock { $entity.ArcLength }
                }
                elseif ($type -eq "AcDbPolyline" -or $type -eq "AcDb2dPolyline" -or $type -eq "AcDb3dPolyline") {
                    try {
                        $length = Invoke-WithRetry -ScriptBlock { $entity.Length }
                    } catch {
                        # игнорируем
                    }
                }

                if ($length -ne $null) {
                    if ($dict.ContainsKey($layer)) {
                        $dict[$layer] += $length
                    } else {
                        $dict[$layer] = $length
                    }
                    $count++
                }
            } catch {
                Write-Host "Ошибка при обработке объекта: $($_.Exception.Message)"
            }
        }

        # Заполняем таблицу на вкладке Длины
        $dataGridViewLength.Columns.Clear()
        $dataGridViewLength.Columns.Add("Length", "Длина") | Out-Null
        $dataGridViewLength.Columns.Add("Layer", "Слой") | Out-Null
        $dataGridViewLength.Rows.Clear()

        $global:currentData = @()
        foreach ($key in ($dict.Keys | Sort-Object)) {
            $val = [math]::Round($dict[$key], 4)
            $dataGridViewLength.Rows.Add($val, $key) | Out-Null
            $global:currentData += [PSCustomObject]@{ Length = $val; Layer = $key }
        }

        # Переключаемся на вкладку Длины
        $tabControl.SelectedTab = $tabLength

        $statusLabel.Text = "Обработано объектов: $count, найдено слоёв: $($dict.Count)"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка: $($_.Exception.Message)", "Ошибка")
        $statusLabel.Text = "Ошибка при выполнении"
    } finally {
        if ($conn) { Disconnect-AutoCAD -acad $conn.App -doc $conn.Doc }
    }
}

function Calculate-Area {
    $global:currentColumns = @("Площадь", "Слой")
    $file = $textBoxFile.Text
    if ([string]::IsNullOrWhiteSpace($file)) {
        [System.Windows.Forms.MessageBox]::Show("Выберите DWG файл!", "Ошибка")
        return
    }
    $version = $comboVersion.SelectedItem

    # Показываем индикатор загрузки
    $statusLabel.Text = "Загрузка и извлечение данных..."
    $mainForm.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $conn = Connect-AutoCAD -version $version -filePath $file
        $acad = $conn.App
        $doc = $conn.Doc
        $model = Invoke-WithRetry -ScriptBlock { $doc.ModelSpace }

        $dict = @{}
        $errorOccurred = $false
        $count = 0
        $zeroAreaCount = 0   # счётчик штриховок с нулевой площадью

        foreach ($entity in $model) {
            try {
                $type = Invoke-WithRetry -ScriptBlock { $entity.ObjectName } -MaxRetries 2
                if ($type -eq "AcDbHatch") {
                    $layer = Invoke-WithRetry -ScriptBlock { $entity.Layer }
                    $area = Invoke-WithRetry -ScriptBlock { $entity.Area }
                    
                    # Проверка на нулевую площадь (с учётом погрешности)
                    if ([math]::Abs($area) -le 1e-6) {
                        $zeroAreaCount++
                    }
                    
                    if ($dict.ContainsKey($layer)) {
                        $dict[$layer] += $area
                    } else {
                        $dict[$layer] = $area
                    }
                    $count++
                }
            } catch {
                $errorOccurred = $true
                Write-Host "Ошибка при обработке штриховки: $($_.Exception.Message)"
            }
        }

        # Заполняем таблицу на вкладке Площади
        $dataGridViewArea.Columns.Clear()
        $dataGridViewArea.Columns.Add("Area", "Площадь") | Out-Null
        $dataGridViewArea.Columns.Add("Layer", "Слой") | Out-Null
        $dataGridViewArea.Rows.Clear()

        $global:currentData = @()
        foreach ($key in ($dict.Keys | Sort-Object)) {
            $val = [math]::Round($dict[$key], 4)
            $dataGridViewArea.Rows.Add($val, $key) | Out-Null
            $global:currentData += [PSCustomObject]@{ Area = $val; Layer = $key }
        }

        # Переключаемся на вкладку Площади
        $tabControl.SelectedTab = $tabArea

        # Выводим предупреждения
        if ($errorOccurred) {
            [System.Windows.Forms.MessageBox]::Show("Обнаружены ошибки при обработке некоторых штриховок.", "Предупреждение")
        }
        if ($zeroAreaCount -gt 0) {
            [System.Windows.Forms.MessageBox]::Show("Обнаружено $zeroAreaCount штриховок с нулевой площадью.", "Предупреждение")
        }
        
        $statusLabel.Text = "Обработано штриховок: $count, найдено слоёв: $($dict.Count). Нулевых: $zeroAreaCount"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка: $($_.Exception.Message)", "Ошибка")
        $statusLabel.Text = "Ошибка при выполнении"
    } finally {
        if ($conn) { Disconnect-AutoCAD -acad $conn.App -doc $conn.Doc }
    }
}

function Calculate-Arc {
    $global:currentColumns = @("Сумма", "Радиус", "Слой")
    $file = $textBoxFile.Text
    if ([string]::IsNullOrWhiteSpace($file)) {
        [System.Windows.Forms.MessageBox]::Show("Выберите DWG файл!", "Ошибка")
        return
    }
    $version = $comboVersion.SelectedItem

    # Показываем индикатор загрузки
    $statusLabel.Text = "Загрузка и извлечение данных..."
    $mainForm.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $conn = Connect-AutoCAD -version $version -filePath $file
        $acad = $conn.App
        $doc = $conn.Doc
        $model = Invoke-WithRetry -ScriptBlock { $doc.ModelSpace }

        # Шаг 1: взорвать все полилинии
        $statusLabel.Text = "Взрыв полилиний..."
        $mainForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()

        $polylines = @()
        foreach ($entity in $model) {
            try {
                $type = Invoke-WithRetry -ScriptBlock { $entity.ObjectName } -MaxRetries 2
                if ($type -eq "AcDbPolyline" -or $type -eq "AcDb2dPolyline" -or $type -eq "AcDb3dPolyline") {
                    $polylines += $entity
                }
            } catch {
                Write-Host "Ошибка при сборе полилиний: $($_.Exception.Message)"
            }
        }

        $explodedCount = 0
        foreach ($pline in $polylines) {
            try {
                $newObjs = Invoke-WithRetry -ScriptBlock { $pline.Explode() } -MaxRetries 2
                $explodedCount++
                [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($pline) | Out-Null
                if ($newObjs) {
                    foreach ($obj in $newObjs) {
                        [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($obj) | Out-Null
                    }
                }
            } catch {
                Write-Host "Ошибка при взрыве полилинии: $($_.Exception.Message)"
            }
        }

        try {
            $doc.Regen(1)  # acAllViewports
        } catch {
            Write-Host "Ошибка при регенерации: $($_.Exception.Message)"
        }
        Start-Sleep -Milliseconds 2000

        $statusLabel.Text = "Сбор данных о дугах..."
        $mainForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()

        $model = Invoke-WithRetry -ScriptBlock { $doc.ModelSpace }

        $dict = @{}
        $count = 0
        $arcCount = 0

        foreach ($entity in $model) {
            try {
                $type = Invoke-WithRetry -ScriptBlock { $entity.ObjectName } -MaxRetries 2
                if ($type -eq "AcDbArc") {
                    $arcCount++
                    $layer = Invoke-WithRetry -ScriptBlock { $entity.Layer }
                    $radius = Invoke-WithRetry -ScriptBlock { $entity.Radius }
                    $len = Invoke-WithRetry -ScriptBlock { $entity.ArcLength }
                    $range = Get-RadiusRange -radius $radius
                    if ($range) {
                        $key = "$layer|$range"
                        if ($dict.ContainsKey($key)) {
                            $dict[$key] += $len
                        } else {
                            $dict[$key] = $len
                        }
                        $count++
                    }
                }
            } catch {
                Write-Host "Ошибка при обработке дуги: $($_.Exception.Message)"
            }
        }

        $rows = @()
        foreach ($key in $dict.Keys) {
            $parts = $key.Split('|')
            $rows += [PSCustomObject]@{
                Layer = $parts[0]
                Radius = $parts[1]
                Length = [math]::Round($dict[$key], 4)
            }
        }
        $sorted = $rows | Sort-Object Layer, { [int]($_.Radius -replace '\D','') }

        # Заполняем таблицу на вкладке Дуги
        $dataGridViewArc.Columns.Clear()
        $dataGridViewArc.Columns.Add("Length", "Сумма") | Out-Null
        $dataGridViewArc.Columns.Add("Radius", "Радиус") | Out-Null
        $dataGridViewArc.Columns.Add("Layer", "Слой") | Out-Null
        $dataGridViewArc.Rows.Clear()

        $global:currentData = $sorted
        foreach ($row in $sorted) {
            $dataGridViewArc.Rows.Add($row.Length, $row.Radius, $row.Layer) | Out-Null
        }

        # Переключаемся на вкладку Дуги
        $tabControl.SelectedTab = $tabArc

        $statusLabel.Text = "Обработано дуг: $count, уникальных групп: $($dict.Count) (взорвано полилиний: $explodedCount)"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка: $($_.Exception.Message)", "Ошибка")
        $statusLabel.Text = "Ошибка при выполнении"
    } finally {
        if ($conn) { Disconnect-AutoCAD -acad $conn.App -doc $conn.Doc }
    }
}

function Calculate-Block {
    $global:currentColumns = @("Кол-во", "Блок (состояние)", "Слой")
    $file = $textBoxFile.Text
    if ([string]::IsNullOrWhiteSpace($file)) {
        [System.Windows.Forms.MessageBox]::Show("Выберите DWG файл!", "Ошибка")
        return
    }
    $version = $comboVersion.SelectedItem

    # Шаг 1: извлечение списка слоёв
    $statusLabel.Text = "Извлечение слоёв..."
    $mainForm.Refresh()
    [System.Windows.Forms.Application]::DoEvents()

    try {
        $conn = Connect-AutoCAD -version $version -filePath $file
        $acad = $conn.App
        $doc = $conn.Doc
        $model = Invoke-WithRetry -ScriptBlock { $doc.ModelSpace }

        $allLayers = @{}
        foreach ($entity in $model) {
            try {
                if ((Invoke-WithRetry -ScriptBlock { $entity.ObjectName } -MaxRetries 2) -eq "AcDbBlockReference") {
                    $layer = Invoke-WithRetry -ScriptBlock { $entity.Layer }
                    $allLayers[$layer] = $true
                }
            } catch {
                Write-Host "Ошибка при сборе слоёв: $($_.Exception.Message)"
            }
        }
        $layerList = $allLayers.Keys | Sort-Object

        $selectedLayer = Show-LayerSelectionDialog -layerList $layerList
        if ($selectedLayer -eq $null) {
            $statusLabel.Text = "Готов"
            return
        }

        $statusLabel.Text = "Загрузка и извлечение данных..."
        $mainForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents()

        $dict = @{}
        $count = 0

        foreach ($entity in $model) {
            try {
                if ((Invoke-WithRetry -ScriptBlock { $entity.ObjectName } -MaxRetries 2) -eq "AcDbBlockReference") {
                    $blkLayer = Invoke-WithRetry -ScriptBlock { $entity.Layer }
                    if ($selectedLayer -and $blkLayer -ne $selectedLayer) { continue }

                    $blkName = Invoke-WithRetry -ScriptBlock { $entity.EffectiveName }
                    $visState = ""
                    $hasVis = $false

                    if ((Invoke-WithRetry -ScriptBlock { $entity.IsDynamicBlock })) {
                        try {
                            $dynProps = Invoke-WithRetry -ScriptBlock { $entity.GetDynamicBlockProperties() }
                            foreach ($prop in $dynProps) {
                                $propName = $prop.PropertyName
                                if ($propName -like "*ВИДИМОСТЬ*" -or $propName -like "*VISIBILITY*") {
                                    $visState = $prop.Value
                                    $hasVis = $true
                                    break
                                }
                            }
                        } catch {
                            Write-Host "Не удалось получить динамические свойства блока $blkName"
                        }
                    }

                    $key = if ($hasVis) { "$blkName | $visState" } else { $blkName }

                    if (-not $dict.ContainsKey($key)) {
                        $dict[$key] = @{}
                    }
                    if ($dict[$key].ContainsKey($blkLayer)) {
                        $dict[$key][$blkLayer]++
                    } else {
                        $dict[$key][$blkLayer] = 1
                    }
                    $count++
                }
            } catch {
                Write-Host "Ошибка при обработке блока: $($_.Exception.Message)"
            }
        }

        $rows = @()
        $blkKeys = $dict.Keys | Sort-Object
        foreach ($blkKey in $blkKeys) {
            $layerDict = $dict[$blkKey]
            $layers = $layerDict.Keys | Sort-Object
            foreach ($l in $layers) {
                $rows += [PSCustomObject]@{
                    Count = $layerDict[$l]
                    Block = $blkKey
                    Layer = $l
                }
            }
        }

        # Заполняем таблицу на вкладке Блоки
        $dataGridViewBlock.Columns.Clear()
        $dataGridViewBlock.Columns.Add("Count", "Кол-во") | Out-Null
        $dataGridViewBlock.Columns.Add("Block", "Блок (состояние)") | Out-Null
        $dataGridViewBlock.Columns.Add("Layer", "Слой") | Out-Null
        $dataGridViewBlock.Rows.Clear()

        $global:currentData = $rows
        foreach ($row in $rows) {
            $dataGridViewBlock.Rows.Add($row.Count, $row.Block, $row.Layer) | Out-Null
        }

        # Переключаемся на вкладку Блоки
        $tabControl.SelectedTab = $tabBlock

        $statusLabel.Text = "Обработано блоков: $count, уникальных вхождений: $($rows.Count)"
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка: $($_.Exception.Message)", "Ошибка")
        $statusLabel.Text = "Ошибка при выполнении"
    } finally {
        if ($conn) { Disconnect-AutoCAD -acad $conn.App -doc $conn.Doc }
    }
}

# ========== Экспорт ==========

function Export-CsvFile {
    if (-not $global:currentData -or $global:currentData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Нет данных для экспорта", "Информация")
        return
    }
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "CSV файлы (*.csv)|*.csv|Все файлы (*.*)|*.*"
    $saveFileDialog.FileName = "export.csv"
    if ($saveFileDialog.ShowDialog() -eq "OK") {
        $global:currentData | Export-Csv -Path $saveFileDialog.FileName -Encoding UTF8 -NoTypeInformation -Delimiter ";"
        [System.Windows.Forms.MessageBox]::Show("Сохранено в $($saveFileDialog.FileName)", "Готово")
    }
}

function Export-Excel {
    if (-not $global:currentData -or $global:currentData.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("Нет данных для экспорта", "Информация")
        return
    }
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveFileDialog.Filter = "Excel файлы (*.xlsx)|*.xlsx|Все файлы (*.*)|*.*"
    $saveFileDialog.FileName = "export.xlsx"
    if ($saveFileDialog.ShowDialog() -ne "OK") { return }

    $excel = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)

        $properties = $global:currentData[0].PSObject.Properties.Name
        for ($col = 0; $col -lt $properties.Count; $col++) {
            $sheet.Cells.Item(1, $col + 1) = $properties[$col]
        }

        $row = 2
        foreach ($item in $global:currentData) {
            for ($col = 0; $col -lt $properties.Count; $col++) {
                $value = $item.($properties[$col])
                $sheet.Cells.Item($row, $col + 1) = if ($value -is [double]) { [double]$value } else { $value }
            }
            $row++
        }

        $usedRange = $sheet.UsedRange
        $usedRange.EntireColumn.AutoFit() | Out-Null
        $usedRange.Rows.Item(1).Font.Bold = $true
        $usedRange.Rows.Item(1).HorizontalAlignment = -4108  # xlCenter

        $workbook.SaveAs($saveFileDialog.FileName, 51)  # 51 = xlOpenXMLWorkbook
        $workbook.Close()
        $excel.Quit()

        [System.Windows.Forms.MessageBox]::Show("Данные экспортированы в Excel: $($saveFileDialog.FileName)", "Готово")
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Ошибка при экспорте в Excel: $($_.Exception.Message)", "Ошибка")
    } finally {
        if ($excel) {
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
    }
}

# ========== Запуск ==========
[void]$mainForm.ShowDialog()