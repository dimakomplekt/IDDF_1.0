Option Explicit


' =========================================================
' ПАКЕТНАЯ СИНХРОНИЗАЦИЯ ЧЕРТЕЖЕЙ
' =========================================================
'
' Общий поток:
'
' ROOT (папка проекта)
'     ↓
' Получаем список подпапок
'     ↓
' Исключаем архив
'     ↓
' В каждой папке ищем *.idw
'     ↓
' Открываем каждый чертёж
'     ↓
' Синхронизируем с моделью (process_drawing)
'     ↓
' Сохраняем
'     ↓
' Закрываем
'     ↓
' Повторяем для всех файлов
'     ↓
' Показываем итог


Sub Main()

    ' =========================
    ' 1. НАСТРОЙКИ
    ' =========================
    
    ' Имя папки, которую нужно игнорировать
    Dim archive_folder_name As String
    archive_folder_name = "1_Архив"

    ' Корневая директория (где лежит текущий документ)
    Dim root_dir As String
    root_dir = ThisDoc.Path


    ' =========================
    ' 2. СЧЁТЧИК РЕЗУЛЬТАТА
    ' =========================
    
    ' Сколько чертежей успешно обработано
    Dim total_count As Integer
    total_count = 0


    ' =========================
    ' 3. ПОЛУЧАЕМ ПОДПАПКИ
    ' =========================
    
    ' Все директории внутри root
    Dim sub_directories() As String
    sub_directories = System.IO.Directory.GetDirectories(root_dir)

    Dim directory_path As String


    ' =========================
    ' 4. ГЛАВНЫЙ TRY (UI-КОНТРОЛЬ)
    ' =========================
    
    Try

        ' Отключаем обновление интерфейса:
        ' ускоряет массовую обработку и убирает мерцание
        ThisApplication.ScreenUpdating = False


        ' =========================
        ' 5. ОБХОД ВСЕХ ПАПОК
        ' =========================
        
        For Each directory_path In sub_directories

            ' Получаем имя папки
            Dim directory_name As String
            directory_name = System.IO.Path.GetFileNameWithoutExtension(directory_path)

            ' Пропускаем архив
            If directory_name = archive_folder_name Then
                Continue For
            End If


            ' =========================
            ' 6. ПОИСК ЧЕРТЕЖЕЙ (*.idw)
            ' =========================
            
            Dim files() As String
            files = System.IO.Directory.GetFiles(directory_path, "*.idw")

            ' Если в папке нет чертежей — пропускаем
            If files.Length = 0 Then
                Continue For
            End If


            ' =========================
            ' 7. ОБРАБОТКА КАЖДОГО ФАЙЛА
            ' =========================
            
            Dim curr_file_path As String

            For Each curr_file_path In files
                
                ' Документ чертежа
                Dim doc As DrawingDocument
                doc = Nothing

                Try

                    ' =====================
                    ' 7.1 ОТКРЫТИЕ
                    ' =====================
                    
                    ' False → открываем в фоне (быстрее)
                    doc = ThisApplication.Documents.Open(curr_file_path, False)


                    ' =====================
                    ' 7.2 СИНХРОНИЗАЦИЯ
                    ' =====================
                    
                    ' Вся логика внутри отдельной функции
                    process_drawing(doc)


                    ' =====================
                    ' 7.3 СОХРАНЕНИЕ
                    ' =====================
                    
                    doc.Save

                    total_count = total_count + 1


                Catch ex As Exception

                    ' Ошибка по конкретному файлу
                    MessageBox.Show("Ошибка: " & curr_file_path & vbCrLf & ex.Message)

                Finally

                    ' =====================
                    ' 7.4 ЗАКРЫТИЕ
                    ' =====================
                    
                    ' False → не сохраняем повторно
                    If Not doc Is Nothing Then doc.Close(False)

                End Try
                
            Next

        Next

    Finally

        ' ВАЖНО: возвращаем обновление UI
        ThisApplication.ScreenUpdating = True

    End Try


    ' =========================
    ' 8. ФИНАЛ
    ' =========================
    
    MessageBox.Show("Готово. Обновлено чертежей: " & total_count)

End Sub


' =========================================================
' СИНХРОНИЗАЦИЯ: ЧЕРТЕЖ ← МОДЕЛЬ
' =========================================================
'
' Поток:
'
' Чертёж
'     ↓
' Берём первый вид
'     ↓
' Получаем модель
'     ↓
' Читаем параметры
'     ↓
' Чистим строки
'     ↓
' Формируем логику (тип документа и т.д.)
'     ↓
' Записываем в iProperties


Sub process_drawing(opened_drawing_document As DrawingDocument)

    ' =========================
    ' 1. ПОЛУЧЕНИЕ МОДЕЛИ
    ' =========================

    Dim opened_sheet As Sheet
    opened_sheet = opened_drawing_document.ActiveSheet

    ' Если нет видов — дальше идти бессмысленно
    If opened_sheet.DrawingViews.Count = 0 Then Exit Sub

    ' Берём первый (базовый) вид
    Dim basic_view As DrawingView
    basic_view = opened_sheet.DrawingViews(1)


    Dim main_model_document As Document

    Try

        ' Получаем модель, связанную с видом
        main_model_document = basic_view.ReferencedDocumentDescriptor.ReferencedDocument

    Catch

        ' Ссылка может быть битой или модель не загружена
        MessageBox.Show("Не удалось получить модель из вида")
        Exit Sub

    End Try

    ' Дополнительная защита
    If main_model_document Is Nothing Then

        MessageBox.Show("Модель не найдена или не загружена")
        Exit Sub

    End If



    ' =========================
    ' 2. ЧТЕНИЕ ПАРАМЕТРОВ
    ' =========================
    
    ' Все значения проходят через:
    ' get_parameters → безопасное чтение
    ' clean_string   → очистка строки

    Dim part_number As String = clean_string(get_parameters(main_model_document, "part_number"))
    Dim part_name As String = clean_string(get_parameters(main_model_document, "part_name"))

    Dim part_type As String = clean_string(get_parameters(main_model_document, "part_type"))

    Dim part_developer As String = clean_string(get_parameters(main_model_document, "part_developer"))

    Dim developer_date As String = clean_string(get_parameters(main_model_document, "developer_date"))
    Dim part_test As String = clean_string(get_parameters(main_model_document, "part_test"))
    Dim test_date As String = clean_string(get_parameters(main_model_document, "test_date"))

    Dim part_tech_control As String = clean_string(get_parameters(main_model_document, "part_tech_control"))
    Dim tech_control_date As String = clean_string(get_parameters(main_model_document, "tech_control_date"))

    Dim part_department_head As String = clean_string(get_parameters(main_model_document, "part_department_head"))
    Dim department_head_date As String = clean_string(get_parameters(main_model_document, "department_head_date"))

    Dim part_norms_control As String = clean_string(get_parameters(main_model_document, "part_norms_control"))
    Dim norms_control_date As String = clean_string(get_parameters(main_model_document, "norms_control_date"))

    Dim part_approved_by As String = clean_string(get_parameters(main_model_document, "part_approved_by"))
    Dim part_approved_date As String = clean_string(get_parameters(main_model_document, "part_approved_date"))

    Dim part_company As String = get_parameters(main_model_document, "part_company")


    ' =========================
    ' 3. ЛОГИКА ДОКУМЕНТА
    ' =========================
    
    Dim document_type As String = ""

    ' Если это сборка — модифицируем номер и тип
    If part_type = "Сборка" Then 
        
        ' Защита от повторного добавления "СБ"
        If Not part_number.EndsWith("СБ") Then
            part_number = part_number & "СБ"
        End If

        document_type = "Сборочный чертеж"

    End If


    ' =========================
    ' 4. ЗАПИСЬ В iPROPERTIES
    ' =========================
    
    ' Всё пишется через safe_properties_setter:
    ' - защита от ошибок
    ' - нормализация значений
    ' - Inventor не падает

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

    ' Тип документа (если определён)
    If document_type <> "" Then
        safe_properties_setter(opened_drawing_document, "Свойства ГОСТ", "Тип документа", document_type)
    End If

    ' Компания
    safe_properties_setter(opened_drawing_document, "Inventor Document Summary Information", "Company", part_company)

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
