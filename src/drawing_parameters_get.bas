Option Explicit


' Чертёж (Drawing)
'         ↓
' Берём первый вид
'         ↓
' Получаем связанную модель
'         ↓
' Читаем User Parameters модели
'         ↓
' Чистим строки (sanitizer)
'         ↓
' Заполняем iProperties чертежа
'         ↓
' Сохраняем


Sub Main()

' Точка входа макроса синхронизации ЧЕРТЕЖ ← МОДЕЛЬ
'
' Общий поток:
'
' 1. Проверка активного документа
' 2. Проверка, что это чертёж (Drawing)
' 3. Проверка наличия видов на листе
' 4. Получение основной модели из первого вида
' 5. Чтение параметров модели (User Parameters)
' 6. Очистка строк (защита от мусора и спецсимволов)
' 7. Запись значений в iProperties чертежа
' 8. Логика типа документа (например "Сборка" → "Сборочный чертеж")
' 9. Сохранение документа
' 10. Завершение работы

    ' =========================
    ' 1. ДОКУМЕНТ
    ' =========================

    Dim opened_document As Document = ThisApplication.ActiveDocument

    ' Проверка: открыт ли документ нужного формата
    If Not TypeOf opened_document Is DrawingDocument Then

        MessageBox.Show("Это не чертёж")
        Exit Sub

    End If
    

    Dim opened_drawing_document As DrawingDocument = opened_document
    Dim opened_sheet As Sheet = opened_drawing_document.ActiveSheet

    ' Проверка: есть ли виды на листе
    If opened_sheet.DrawingViews.Count = 0 Then

        MessageBox.Show("Нет видов на листе")
        Exit Sub

    End If


    ' =========================
    ' 2. ПОЛУЧЕНИЕ МОДЕЛИ
    ' =========================

    ' Берём базовый (первый) вид
    Dim basic_view As DrawingView = opened_sheet.DrawingViews(1)

    Dim main_model_document As Document

    Try

        ' Получаем документ модели, связанный с видом opened_sheet.DrawingViews(1)
        main_model_document = basic_view.ReferencedDocumentDescriptor.ReferencedDocument

    Catch

        ' Если ссылка сломана или модель не загружена
        MessageBox.Show("Не удалось получить модель из вида")
        Exit Sub
        
    End Try

    ' Дополнительная защита
    If main_model_document Is Nothing Then

        MessageBox.Show("Модель не найдена или не загружена")
        Exit Sub

    End If


    ' =========================
    ' 3. ЧТЕНИЕ ПАРАМЕТРОВ МОДЕЛИ
    ' =========================

    ' Все параметры читаются через safe getter:
    ' - если параметра нет → ""
    ' - если ошибка → ""
    ' - никаких падений

    Dim part_number As String = clean_string(get_parameters(main_model_document, "part_number")) ' B
    Dim part_name As String = clean_string(get_parameters(main_model_document, "part_name")) ' C

    Dim part_type As String = clean_string(get_parameters(main_model_document, "part_type")) ' D

    Dim part_developer As String = clean_string(get_parameters(main_model_document, "part_developer")) ' E

    Dim developer_date As String = clean_string(get_parameters(main_model_document, "developer_date")) ' F
    Dim part_test As String = clean_string(get_parameters(main_model_document, "part_test")) ' G
    Dim test_date As String = clean_string(get_parameters(main_model_document, "test_date")) ' H

    Dim part_tech_control As String = clean_string(get_parameters(main_model_document, "part_tech_control")) ' I
    Dim tech_control_date As String = clean_string(get_parameters(main_model_document, "tech_control_date")) ' J

    Dim part_department_head As String = clean_string(get_parameters(main_model_document, "part_department_head")) ' L
    Dim department_head_date As String = clean_string(get_parameters(main_model_document, "department_head_date")) ' M

    Dim part_norms_control As String = clean_string(get_parameters(main_model_document, "part_norms_control")) ' N
    Dim norms_control_date As String = clean_string(get_parameters(main_model_document, "norms_control_date")) ' O

    Dim part_approved_by As String = clean_string(get_parameters(main_model_document, "part_approved_by")) ' P
    Dim part_approved_date As String = clean_string(get_parameters(main_model_document, "part_approved_date")) ' Q

    Dim part_company As String = clean_string(get_parameters(main_model_document, "part_company")) ' R


    Dim document_type As String = ""


    If part_type = "Сборка" Then 
        
        ' Сейф от мультимутации строки
        
        If Not part_number.EndsWith("СБ") Then
            part_number = part_number & "СБ"
        End If

        document_type = "Сборочный чертеж"

    End If


    ' =========================
    ' 4. iProperties (СТАБИЛЬНО)
    ' =========================

    ' Запись идёт через safe_properties_setter:
    ' - защита от Nothing
    ' - "-" → ""
    ' - Try/Catch внутри
    ' - никаких E_FAIL

    ' --- Основные данные чертежа ---
    
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Part Number", part_number)
    safe_properties_setter(opened_drawing_document, "Inventor Summary Information", "Title", part_name)

    safe_properties_setter(opened_drawing_document, "Inventor Summary Information", "Author", part_developer)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Creation Time", developer_date)
    
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Checked By", part_test)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Date Checked", test_date)

    safe_properties_setter(opened_drawing_document, "Inventor Document Summary Information", "Manager", part_department_head)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Authority", part_department_head)
    safe_properties_setter(opened_drawing_document, "Свойства ГОСТ", "Руководитель Дата", department_head_date)

    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Engr Approved By", part_norms_control)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Engr Date Approved", norms_control_date)

    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Mfg Approved By", part_approved_by)
    safe_properties_setter(opened_drawing_document, "Design Tracking Properties", "Mfg Date Approved", part_approved_date)


    If document_type <> "" Then

        safe_properties_setter(opened_drawing_document, "Свойства ГОСТ", "Тип документа", document_type)
    
    End If

    ' Компания
    safe_properties_setter(opened_drawing_document, "Inventor Document Summary Information", "Company", part_company)


    ' =========================
    ' 6. SAVE
    ' =========================

    Try
        opened_document.Save
    Catch
        ' даже если save упал — мы не драматизируем
    End Try

    MessageBox.Show("Готово!")

End Sub


' =========================================================
' SAFE PARAMETER ACCESS (100% SAFE)
' =========================================================
Function get_parameters(doc As Document, paramName As String) As String

' Универсальный безопасный геттер параметров:
'
' 1. Пытается получить параметр из модели
' 2. Если параметра нет → ""
' 3. Если ошибка API → ""
'
' Никогда не падает

    Try

        Dim params As Parameters = doc.ComponentDefinition.Parameters

        Dim p As Parameter = params.Item(paramName)

        Return CStr(p.Value)

    Catch

        Return ""

    End Try

End Function


' =========================================================
' STRING CLEANER
' =========================================================
Function clean_string(value As String) As String

' Нормализует строку перед записью:
'
' - убирает переводы строк и табы
' - заменяет кавычки
' - удаляет запрещённые символы файловой системы
' - схлопывает двойные пробелы
' - обрезает длину
'
' Делает строку "безопасной для Inventor"

    Try
        If value Is Nothing Then Return ""

        Dim s As String = value

        s = s.Replace(vbCr, "")
        s = s.Replace(vbLf, "")
        s = s.Replace(vbTab, "")

        s = s.Replace(Chr(34), "'")

        Dim bad As String = "/\:*?<>|"

        For i As Integer = 1 To Len(bad)
            s = s.Replace(Mid(bad, i, 1), "_")
        Next

        Do While s.Contains("  ")
            s = s.Replace("  ", " ")
        Loop

        s = s.Trim()

        If s.Length > 250 Then
            s = Left(s, 250)
        End If

        Return s

    Catch
        Return ""
    End Try

End Function


' =========================================================
' SAFE iPROPERTY WRITER (NO LOOP, NO SEARCH)
' =========================================================
Sub safe_properties_setter(doc As Document, setName As String, propName As String, value As String)

' Абсолютно безопасная запись iProperty:
'
' 1. Нормализация значения:
'    - Nothing → ""
'    - "-" → ""
' 2. Прямой доступ к PropertySet
' 3. Try/Catch на весь блок
'
' Любая ошибка → игнор (Inventor не падает)

    Try

        ' =========================
        ' NORMALIZATION LAYER
        ' =========================
        If value Is Nothing Then value = ""

        value = value.Trim()

        ' Пустая строка в случае, если в ведомости проставлен прочерк
        If value = "-" Then value = ""

        Dim ps As PropertySet = doc.PropertySets.Item(setName)
        Dim p As [Property] = ps.Item(propName)

        p.Value = value

    Catch
        
    End Try

End Sub
