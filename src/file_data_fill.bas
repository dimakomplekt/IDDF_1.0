Option Explicit


' Сборка (Assembly)
'         ↓
' Рекурсивный обход всех компонентов
'         ↓
' По имени ищем строку в Excel
'         ↓
' Достаём значения
'         ↓
' Пишем:
'     - User Parameters
'     - iProperties
'         ↓
' Логируем


Sub Main()

' Точка входа всей системы синхронизации Inventor ↔ Excel
'
' Общий поток:
'
' 1. Проверка активного документа
' 2. Проверка, что это сборка (Assembly)
' 3. Определение корневой папки проекта
' 4. Поиск Excel-ведомости
' 5. Инициализация debug-лога
' 6. Запуск Excel через COM
' 7. Загрузка таблицы
' 8. Инициализация visited (защита от повторной обработки)
' 9. Обход корневого документа и всей структуры сборки
' 10. Закрытие Excel
' 11. Запись итогового лога
' 12. Завершение работы

    ' Имя листа части ведомости для парсинга
    Dim SHEET_FOR_DATA_PARSE = "Ведомость для парсинга" 

    Try

        Dim active_document As Document
        active_document = ThisApplication.ActiveDocument

        ' Проверка: есть ли открытый документ
        If active_document Is Nothing Then
            MessageBox.Show("Откройте сборку")
            Exit Sub
        End If

        ' Проверка: это именно сборка, а не деталь (мы обязаны запускать макрос из общей сборки изделия)
        If Not TypeOf active_document Is AssemblyDocument Then
            MessageBox.Show("Нужна сборка")
            Exit Sub
        End If

        Dim root_assembly As AssemblyDocument
        root_assembly = active_document


        ' Базовая папка проекта (где лежит сборка)
        Dim root_folder As String
        root_folder = get_parent_folder_name(root_assembly.FullFileName)


        ' Путь к ведомости состава изделия
        Dim excel_path As String
        excel_path = path_combine(root_folder, "Ведомость состава изделия.xlsx")

        ' Проверка наличия Excel-файла
        If Not System.IO.File.Exists(excel_path) Then
            MessageBox.Show("Excel не найден")
            Exit Sub
        End If

        ' Файл отладки (лог ошибок и пропусков)
        Dim debug_path As String
        debug_path = path_combine(root_folder, "debug.log")

        ' Инициализация debug-лога
        write_text_file(debug_path, "START " & Now & vbCrLf)

        ' Запуск Excel через COM (late binding)
        Dim excel_app As Object
        excel_app = CreateObject("Excel.Application")

        ' Ведомость
        Dim workbook As Object
        Dim worksheet As Object

        ' Открытие ведомости
        workbook = excel_app.Workbooks.Open(excel_path)
        worksheet = get_worksheet_by_name_safe(workbook, SHEET_FOR_DATA_PARSE)

        ' Список уже обработанных файлов (защита от повторов)
        Dim visited As Object
        visited = CreateObject("Scripting.Dictionary")


        ' Итоговый лог обработки всех документов
        Dim log As String

        ' Обработка корневой сборки
        process_document(root_assembly, worksheet, visited, log, True, debug_path)
        ' Рекурсивный обход всей структуры сборки
        process_occurrences(root_assembly.ComponentDefinition.Occurrences, worksheet, visited, log, debug_path)

        ' Закрытие Excel без сохранения изменений
        workbook.Close(False)
        excel_app.Quit


        ' Запись итогового отчёта
        Dim log_path As String
        log_path = write_log_file(root_folder, log)

        MessageBox.Show("Готово: " & log_path)


    Catch ex As Exception
        ' Фатальная ошибка всего процесса
        MessageBox.Show("Ошибка: " & ex.Message)
    End Try

End Sub


' =========================
' OCCURRENCES
' =========================

' Рекурсивно обходит все вхождения сборки и обрабатывает документы
Sub process_occurrences(occurrences As Object, worksheet As Object, visited As Object, ByRef log As String, debug_path As String)

' Что делает:
'
' 1. Перебирает все вхождения (occurrences)
' 2. Пытается получить документ каждого вхождения
' 3. Пропускает, если документа нет
' 4. Проверяет, был ли документ уже обработан (через мапу visited, выданную для проверки)
' 5. Если нет — обрабатывает через process_document
' 6. Рекурсивно спускается во вложенные вхождения (SubOccurrences)
'
' Особенности:
'
' - visited защищает от повторной обработки одинаковых файлов
' - работает рекурсивно (обход дерева любой глубины)

    Dim occ As Object

    For Each occ In occurrences

        Try

            Dim occ_document As Document
            occ_document = Nothing

            Try
                ' Получение документа, связанного с вхождением
                occ_document = occ.Definition.Document
            Catch
                ' Если получить документ не удалось — пропускаем
                append_debug(debug_path, "SKIP OCC (no doc)")
            End Try

            ' Если документа нет — переходим к следующему
            If occ_document Is Nothing Then
                Continue For
            End If

            Dim key As String
            ' Уникальный ключ документа (полный путь)
            key = occ_document.FullFileName

            ' Проверка: обрабатывался ли уже этот документ
            If Not visited.Exists(key) Then
                ' Помечаем как обработанный
                visited.Add(key, True)

                ' Обработка документа
                process_document(occ_document, worksheet, visited, log, False, debug_path)
            End If

            ' Если есть вложенные элементы — уходим глубже
            If occ.SubOccurrences.Count > 0 Then
                process_occurrences(occ.SubOccurrences, worksheet, visited, log, debug_path)
            End If

        Catch ex As Exception
            ' Логирование ошибки обработки вхождения
            append_debug(debug_path, "OCC ERR " & ex.Message)
        End Try

    Next

End Sub


' =========================
' DOCUMENT
' =========================

' Обрабатывает один документ Inventor на основе данных из Excel
Sub process_document(doc As Document, worksheet As Object, visited As Object, ByRef log As String, is_root As Boolean, debug_path As String)

' Что происходит внутри:
'
' 1. Нормализация имени документа
' 2. Поиск соответствующей строки в Excel
' 3. Проверка статуса (фильтр "закуп")
' 4. Извлечение значений из строки
' 5. Применение данных:
'    - пользовательские параметры
'    - iProperties
' 6. Добавление записи в лог
' 7. Сохранение документа
' 8. Закрытие (если не корневой)

    Try

        Dim name As String
        ' Приведение имени файла к нормализованному виду
        name = normalize_name(doc.DisplayName)

        Dim row As Long
        ' Поиск строки в таблице Excel по имени
        row = find_table_row_by_name(worksheet, name)

        ' Если строка не найдена — прекращаем обработку
        If row = 0 Then Exit Sub

        Dim status As String
        ' Получение статуса (например: "закуп", "собств.")
        status = get_cell_text_by_row_col(worksheet, row, 6)

        ' Пропуск покупных изделий (с приведением к нижнему регистру)
        If LCase(status) = "закуп" Then Exit Sub

        Dim values As Object
        ' Извлечение всех значений строки в словарь вспомогательной функцией
        values = get_row_values(worksheet, row)

        ' Применение параметров к документу
        set_user_parameters(doc, values, debug_path)
        ' Заполнение iProperties документа
        set_IProperties(doc, values, debug_path)

        ' Добавление информации в лог обработки
        log = log & build_log_section(name, values)

        ' Сохранение документа
        doc.Save

        ' Закрытие документа, если это не корневой (верхний) документ
        If Not is_root Then
            doc.Close(True)
        End If

    Catch ex As Exception
        ' Логирование ошибки обработки документа
        append_debug(debug_path, "DOC ERR " & ex.Message)
    End Try

End Sub



' =========================
' EXCEL
' =========================

' Базовые функции работы с Excel через COM API.
' Особенности:
' - Работа идёт через Object (late binding) → нет строгой типизации
' - Любой доступ к ячейке может выбросить исключение
' - Значения могут быть Nothing, числом, датой или строкой
'
' Поэтому:
' - всё оборачивается в Try/Catch
' - любые значения приводятся к строке
' - любые ошибки → безопасный возврат ""

' Геттер текстового значения ячейки по её номеру
Function get_cell_text_by_row_col(ws As Object, row As Long, col As Long) As String

    Try
        Dim v As Object

        ' Чтение значения ячейки через COM (может вернуть любой тип или Nothing)
        v = ws.Cells(row, col).Value

         ' Excel часто возвращает Nothing для пустых ячеек → нормализуем в пустую строку
        If v Is Nothing Then
            Return ""
        End If

        ' Приведение к строке + удаление лишних пробелов
        Return CStr(v).Trim()

    Catch
        ' Любая ошибка доступа к Excel → считаем ячейку пустой
        Return ""
    End Try

End Function


' Поиск строки в Excel по имени (колонка B)
' Используется нормализация имени, чтобы сопоставить:
'   - имена файлов Inventor
'   - значения из Excel
Function find_table_row_by_name(ws As Object, searchName As String) As Long

    Dim row As Long
    row = 2 ' начинаем после заголовка таблицы

    Dim blank_count As Long
    blank_count = 0 ' счётчик подряд идущих пустых строк (для определения конца таблицы)

    Dim normalized_search As String
    normalized_search = normalize_name(searchName)

    ' Ограничение поиска (защита от бесконечного цикла при "грязном" Excel)
    ' MAGIC NUMBER - MAX TABLE SIZE !!!
    While row <= 2000

        Dim cell_value As String

        ' Получаем значение ячейки через геттер
        cell_value = get_cell_text_by_row_col(ws, row, 2)

        ' Если строка пустая — увеличиваем счётчик
        If cell_value = "" Then
            blank_count += 1
        Else
            blank_count = 0
        End If

        ' Если встретили длинную "пустоту" — считаем, что таблица закончилась
        If blank_count > 20 Then Exit While

        If cell_value <> "" Then

            Dim normalized_cell As String
            normalized_cell = normalize_name(cell_value)

            ' Сравнение без учёта регистра (доп. защита через normalize)
            If StrComp(normalized_cell, normalized_search, vbTextCompare) = 0 Then
                Return row
            End If

        End If

        row += 1

    End While

    ' Если не нашли — возвращаем 0 (сигнал "нет строки")
    Return 0

End Function


' Геттер значений для записи в параметры модели
Function get_row_values(ws As Object, row As Long) As Object

    ' Формирует словарь значений из строки Excel:
    ' каждая колонка мапится в именованное поле (ключ),
    ' чтобы дальше работать не с индексами, а с осмысленными данными

    Dim d As Object
    d = CreateObject("Scripting.Dictionary")

    ' Чтение значений из фиксированных колонок таблицы
    ' (жёсткая привязка структуры Excel → структуры данных!!!)

    d.Add("position_number", get_cell_text_by_row_col(ws, row, 1))
    d.Add("part_number", get_cell_text_by_row_col(ws, row, 3))
    d.Add("part_name", get_cell_text_by_row_col(ws, row, 4))
    d.Add("part_type", get_cell_text_by_row_col(ws, row, 5))
    d.Add("part_developer", get_cell_text_by_row_col(ws, row, 7))
    d.Add("developer_date", get_cell_text_by_row_col(ws, row, 8))
    d.Add("part_test", get_cell_text_by_row_col(ws, row, 9))
    d.Add("test_date", get_cell_text_by_row_col(ws, row, 10))
    d.Add("part_tech_control", get_cell_text_by_row_col(ws, row, 11))
    d.Add("tech_control_date", get_cell_text_by_row_col(ws, row, 12))
    d.Add("part_department_head", get_cell_text_by_row_col(ws, row, 13))
    d.Add("department_head_date", get_cell_text_by_row_col(ws, row, 14))
    d.Add("part_norms_control", get_cell_text_by_row_col(ws, row, 15))
    d.Add("norms_control_date", get_cell_text_by_row_col(ws, row, 16))
    d.Add("part_approved_by", get_cell_text_by_row_col(ws, row, 17))
    d.Add("part_approved_date", get_cell_text_by_row_col(ws, row, 18))
    d.Add("part_company", get_cell_text_by_row_col(ws, row, 19))

    ' Возвращаем "плоский объект данных" для дальнейшей передачи
    ' (в параметры, iProperties и лог)
    
    get_row_values = d

End Function


' =========================
' NORMALIZE
' =========================
Function normalize_name(s As String) As String

    ' Нормализует имя для корректного сравнения:
    ' - приводит к нижнему регистру
    ' - отбрасывает путь (если передан полный путь)
    ' - убирает лишние пробелы
    '
    ' Используется для сопоставления:
    ' имя файла Inventor ↔ значение в Excel

    Dim name As String

    name = LCase(s)

    ' Если передан полный путь — оставляем только имя файла
    name = System.IO.Path.GetFileName(name)

    normalize_name = Trim(name)

End Function


Function get_file_name_no_format(path As String) As String

    ' Возвращает имя файла без расширения
    ' (например: "part.ipt" → "part")

    Return System.IO.Path.GetFileNameWithoutExtension(path)

End Function

' Получение корректного имени листа
Function get_worksheet_by_name_safe(workbook As Object, target As String) As Object

    Dim sh As Object

    For Each sh In workbook.Worksheets

        If normalize_name(sh.Name) = normalize_name(target) Then
            Return sh
        End If

    Next

    Return Nothing

End Function


' =========================
' PARAMETERS (FIXED)
' =========================

' Устанавливает набор пользовательских параметров документа из словаря values
Sub set_user_parameters(doc As Document, values As Object, debug_path As String)

    Try

        ' Установка каждого параметра через безопасный метод

        safe_set_param(doc, "position_number", values("position_number"), debug_path)
        safe_set_param(doc, "part_number", values("part_number"), debug_path)
        safe_set_param(doc, "part_name", values("part_name"), debug_path)
        safe_set_param(doc, "part_type", values("part_type"), debug_path)
        safe_set_param(doc, "part_developer", values("part_developer"), debug_path)
        safe_set_param(doc, "developer_date", values("developer_date"), debug_path)
        safe_set_param(doc, "part_test", values("part_test"), debug_path)
        safe_set_param(doc, "test_date", values("test_date"), debug_path)
        safe_set_param(doc, "part_tech_control", values("part_tech_control"), debug_path)
        safe_set_param(doc, "tech_control_date", values("tech_control_date"), debug_path)
        safe_set_param(doc, "part_department_head", values("part_department_head"), debug_path)
        safe_set_param(doc, "department_head_date", values("department_head_date"), debug_path)
        safe_set_param(doc, "part_norms_control", values("part_norms_control"), debug_path)
        safe_set_param(doc, "norms_control_date", values("norms_control_date"), debug_path)
        safe_set_param(doc, "part_approved_by", values("part_approved_by"), debug_path)
        safe_set_param(doc, "part_approved_date", values("part_approved_date"), debug_path)
        safe_set_param(doc, "part_company", values("part_company"), debug_path)

    Catch ex As Exception
        ' Логирование ошибки при установке параметров
        append_debug(debug_path, "current_parameter ERR " & ex.Message)
    End Try

End Sub


' Безопасно устанавливает или создаёт пользовательский параметр
Sub safe_set_param(doc As Document, name As String, value As String, debug_path As String)

    Try

        ' 1) Попытка установить параметр через iLogic (самый прямой путь)
        Try
            Parameter(name) = value
            Exit Sub
        Catch
        End Try

        ' 2) Резервный путь через API Inventor
        Dim user_parameters As Object
        user_parameters = doc.ComponentDefinition.Parameters.UserParameters

        Dim current_parameter As Object
        current_parameter = Nothing

        ' Попытка получить существующий параметр для простой замены в случае существования
        Try
            current_parameter = user_parameters.Item(name)
        Catch
            current_parameter = Nothing
        End Try

        ' Если параметра нет — создаём
        If current_parameter Is Nothing Then
            user_parameters.AddByValue(name, value, UnitsTypeEnum.kTextUnits)
        Else
            ' Если есть — обновляем значение (как текст)
            current_parameter.Expression = """" & value & """"
        End If

    ' Логирование ошибки установки конкретного параметра
    Catch ex As Exception
        append_debug(debug_path, "SET current_parameter ERR " & name & " " & ex.Message)
    End Try

End Sub


' =========================
' IProperties
' =========================

' Заполняет стандартные IProperties документа (метаданные)
Sub set_IProperties(doc As Document, values As Object, debug_path As String)

    Try

        ' === Design Tracking Properties ===
        ' Основные инженерные и производственные атрибуты
        doc.PropertySets.Item("Design Tracking Properties").Item("Stock Number").Value = values("position_number")
        doc.PropertySets.Item("Design Tracking Properties").Item("Part Number").Value = values("part_number")
        doc.PropertySets.Item("Design Tracking Properties").Item("Creation Time").Value = values("developer_date")
        doc.PropertySets.Item("Design Tracking Properties").Item("Checked By").Value = values("part_test")
        doc.PropertySets.Item("Design Tracking Properties").Item("Date Checked").Value = values("test_date")
        doc.PropertySets.Item("Design Tracking Properties").Item("Engr Approved By").Value = values("part_norms_control")
        doc.PropertySets.Item("Design Tracking Properties").Item("Engr Date Approved").Value = values("norms_control_date")
        doc.PropertySets.Item("Design Tracking Properties").Item("Mfg Approved By").Value = values("part_approved_by")
        doc.PropertySets.Item("Design Tracking Properties").Item("Mfg Date Approved").Value = values("part_approved_date")


        ' === Inventor Summary Information ===
        ' Общая информация о документе
        doc.PropertySets.Item("Inventor Summary Information").Item("Title").Value = values("part_name")
        doc.PropertySets.Item("Inventor Summary Information").Item("Author").Value = values("part_developer")

        ' === Inventor Document Summary Information ===
        ' Организационные данные
        doc.PropertySets.Item("Inventor Document Summary Information").Item("Manager").Value = values("part_department_head")
        doc.PropertySets.Item("Design Tracking Properties").Item("Authority").Value = values("part_department_head")
        
        doc.PropertySets.Item("Inventor Document Summary Information").Item("Company").Value = values("part_company")

    Catch ex As Exception
        ' Логирование ошибки при работе с IProperties
        append_debug(debug_path, "IPROP ERR " & ex.Message)
    End Try

End Sub


' =========================
' LOG
' =========================
Function build_log_section(name As String, values As Object) As String

    ' Формирует текстовый блок лога для одного документа:
    ' имя + все извлечённые из Excel значения (как "снимок состояния")
    ' используется для накопления общего отчёта

    Dim s As String

    s = vbCrLf & "---------------------------------------------------------" & vbCrLf
    s = s & name & vbCrLf
    s = s & "---------------------------------------------------------" & vbCrLf

    s = s & vbCrLf & "position_number: " & values("position_number") & vbCrLf
    s = s & "part_number: " & values("part_number") & vbCrLf
    s = s & "part_name: " & values("part_name") & vbCrLf
    s = s & "part_type: " & values("part_type") & vbCrLf

    s = s & "part_developer: " & values("part_developer") & vbCrLf
    s = s & "developer_date: " & values("developer_date") & vbCrLf

    s = s & "part_test: " & values("part_test") & vbCrLf
    s = s & "test_date: " & values("test_date") & vbCrLf

    s = s & "part_tech_control: " & values("part_tech_control") & vbCrLf
    s = s & "tech_control_date: " & values("tech_control_date") & vbCrLf

    s = s & "part_department_head: " & values("part_department_head") & vbCrLf
    s = s & "department_head_date: " & values("department_head_date") & vbCrLf

    s = s & "part_norms_control: " & values("part_norms_control") & vbCrLf
    s = s & "norms_control_date: " & values("norms_control_date") & vbCrLf

    s = s & "part_approved_by: " & values("part_approved_by") & vbCrLf
    s = s & "part_approved_date: " & values("part_approved_date") & vbCrLf

    s = s & "part_company: " & values("part_company") & vbCrLf & vbCrLf 

    s = s & "---------------------------------------------------------" & vbCrLf & vbCrLf

    build_log_section = s

End Function


Function write_log_file(folder As String, log As String) As String

    ' Записывает финальный лог (накопленный по всем документам) в файл log.txt
    ' возвращает полный путь к файлу для отображения пользователю

    Dim path As String
    path = path_combine(folder, "log.txt")

    System.IO.File.WriteAllText(path, log)

    Return path

End Function


Sub write_text_file(path As String, text As String)

    ' Полная перезапись файла (используется для старта debug-лога)
    System.IO.File.WriteAllText(path, text)

End Sub


Sub append_debug(path As String, text As String)

    ' Дописывает строку в debug-лог:
    ' добавляет временную метку для последующего анализа ошибок/пропусков
    System.IO.File.AppendAllText(path, Now & " " & text & vbCrLf)

End Sub



' =========================
' PATH
' =========================

Function get_parent_folder_name(path As String) As String

    ' Возвращает путь к родительской папке указанного файла/пути
    Return System.IO.Path.GetDirectoryName(path)

End Function


Function path_combine(a As String, b As String) As String

    ' Объединяет два пути в один корректный (с учётом разделителей ОС)
    Return System.IO.Path.Combine(a, b)

End Function
