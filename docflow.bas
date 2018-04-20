Attribute VB_Name = "Module1"
Sub docflow()
Attribute docflow.VB_ProcData.VB_Invoke_Func = "q\n14"
Dim value As Variant
Dim val2 As Variant
Dim sFilesPath As String

    'добавить строчку
    Sheets("Вся входящая корреспонденция").Select
    Range("A4").Select
    Selection.EntireRow.Insert
    
        'наименование папки
        Sheets("Вся входящая корреспонденция").Select
        Range("B4").Select
        value = InputBox("'сз'=Служебные записки входящие;" & vbCrLf & "'ст'=Сторонние организации входящее," & vbCrLf & " ** = выход", "НАИМЕНОВАНИЕ ПАПКИ ")
        If value = "**" Then Exit Sub
        If value = "ст" Then
            ActiveCell.value = "Сторонние организации входящее"
        ElseIf value = "сз" Then
            ActiveCell.value = "Служебные записки входящие"
        Else
            ActiveCell.value = value
        End If
        
        
    'Исходящий номер
    Sheets("Вся входящая корреспонденция").Select
    Range("C4").Select
    value = InputBox("№ регистрации" & vbCrLf & "** = выход", "ИСХОДЯЩЕЕ")
    If value = "**" Then Exit Sub
    ActiveCell.value = value

    With Application.FileDialog(msoFileDialogFilePicker)
        
       If .Show = False Then Exit Sub
       .Filters.Add "All Files", "*.*"
       sFilesPath = .SelectedItems(1)
   End With
      ActiveSheet.Hyperlinks.Add Anchor:=Range("C4"), Address:=sFilesPath
    
        'Исходящий дата
        Sheets("Вся входящая корреспонденция").Select
        Range("D4").Select
        value = InputBox("Дата регистрации" & vbCrLf & " 0=сегодня" & vbCrLf & "число= число этого месяца, пример: при вводе 19 будет добавлено 19.04.2018" & vbCrLf & "** = выход", "ИСХОДЯЩЕЕ")
        If value = "**" Then Exit Sub
            '0=сегодня
        If value = 0 Then value = Date
            'одно число = дата этого месяца
        If IsNumeric(value) Then value = DateSerial(Year(Date), Month(Date), value)
        ActiveCell.value = value
    
    'В ответ на номер
    Sheets("Вся входящая корреспонденция").Select
    Range("E4").Select
    value = InputBox("№ регистрации" & vbCrLf & "** = выход", "В ОТВЕТ НА")
    If value = "**" Then Exit Sub
    ActiveCell.value = value
    
        'В ответ на дата
        Sheets("Вся входящая корреспонденция").Select
        Range("F4").Select
        value = InputBox("Дата регистрации" & vbCrLf & "0=сегодня" & vbCrLf & "число= число этого месяца" & vbCrLf & "пример: при вводе 19 будет добавлено 19.04.2018" & vbCrLf & "** = выход", "В ОТВЕТ НА")
        If value = "**" Then Exit Sub
            '0=сегодня
        If value = 0 Then value = Date
            'одно число = дата этого месяца
        If IsNumeric(value) Then value = DateSerial(Year(Date), Month(Date), value)
        ActiveCell.value = value

    'Входящий номер
    Sheets("Вся входящая корреспонденция").Select
    Range("G4").Select
    value = InputBox("№ регистрации" & vbCrLf & "** = выход", "ВХОДЯЩИЙ")
    If value = "**" Then Exit Sub
    ActiveCell.value = value
    
        'Входящий дата
        Sheets("Вся входящая корреспонденция").Select
        Range("H4").Select
        value = InputBox("Дата регистрации" & vbCrLf & "0=сегодня" & vbCrLf & "число= число этого месяца, пример: при вводе 19 будет добавлено 19.04.2018" & vbCrLf & "** = выход", "ВХОДЯЩИЙ")
        If value = "**" Then Exit Sub
            '0=сегодня
        If value = 0 Then value = Date
            'одно число = дата этого месяца
        If IsNumeric(value) Then value = DateSerial(Year(Date), Month(Date), value)
        ActiveCell.value = value

    'Кому (адресат)(в им. падеже)
    Sheets("Вся входящая корреспонденция").Select
    Range("I4").Select
    value = InputBox("Адресат в им. падеже" & vbCrLf & "оп = Андреев" & vbCrLf & "нв = Куприянов" & vbCrLf & "нв+ = Куприянов и добавить еще" & vbCrLf & "** = выход", "КОМУ")
    Select Case value
        Case "**"
            Exit Sub
        Case "оп"
            value = "О.П. Андрееву"
        Case "нв"
            value = "Н.В. Куприянову"
        Case "нв+"
            val2 = InputBox("куприянову/..." & vbCrLf & "ис = Аристархову")
            If val2 = "ис" Then
                value = "Н.В.Куприянову/ И.С. Аристархову"
            Else
                value = "Н.В.Куприянову/ " + val2
            End If
        End Select
    ActiveCell.value = value

        'Кому (Организация)(в им. падеже)
        Sheets("Вся входящая корреспонденция").Select
        Range("J4").Select
        value = InputBox("Организация" & vbCrLf & "адм = Администрация" & vbCrLf & "адм+ = Администрация + еще кто то" & vbCrLf & "** = выход", "КОМУ")
        Select Case value
            Case "**"
                Exit Sub
            Case "адм"
                value = "Администрация"
            Case "адм+"
                val2 = InputBox("Администрация/..." & vbCrLf & "нн = Н.Новгород ф" & vbCrLf & "спб = Петербург ф" & vbCrLf & _
                "мск = Москва ф" & vbCrLf & "мах = Махачкала ф" & vbCrLf & "под = Подольск ф" & vbCrLf & "сар = Саратов ф" & vbCrLf & "ст = Ставрополь ф" & vbCrLf & _
                "тю = Тюмень ф" & vbCrLf & "нск = Новосибирск ф" & vbCrLf & "** = выход")
                Select Case val2
                Case "**"
                    Exit Sub
                Case "нн"
                    val2 = "Нижегородский филиал"
                Case "спб"
                    val2 = "Санкт-Петербургский филиал"
                Case "мск"
                    val2 = "Московский филиал"
                Case "гсг"
                    val2 = "АО Гипроспецгаз"
                Case "ги"
                    val2 = "ООО Газпром инвест"
                Case "мах"
                    val2 = "Махачкалинский филиал"
                Case "под"
                    val2 = "Подольский филиал"
                Case "сар"
                    val2 = "Саратовский филиал"
                Case "тю"
                    val2 = "Тюменский филиал"
                Case "нск"
                    val2 = "Новосибирский филиал"
                Case "ст"
                    val2 = "Ставропольский филиал"
                End Select
                value = "Администрация/ " + val2
            End Select
        ActiveCell.value = value
    
    'От кого - Организация
    Sheets("Вся входящая корреспонденция").Select
    Range("L4").Select
    value = InputBox("От кого организация" & vbCrLf & "нн = Н.Новгород ф" & vbCrLf & _
        "спб = Петербург ф" & vbCrLf & "мск = Москва ф" & vbCrLf & "мах = Махачкала ф" & vbCrLf & _
        "под = Подольск ф" & vbCrLf & "сар = Саратов ф" & vbCrLf & "ст = Ставрополь ф" & vbCrLf & _
        "тю = Тюмень ф" & vbCrLf & "нск = Новосибирск ф" & vbCrLf & "адм = Администрация" & vbCrLf & _
        "--------------" & vbCrLf & "ги = Газпром инвест" & vbCrLf & "гсг = Гипроспецгаз" & vbCrLf & "** = выход", "ОТ КОГО")
    Select Case value
        Case "**"
            Exit Sub
        Case "нн"
            value = "Нижегородский филиал"
        Case "спб"
            value = "Санкт-Петербургский филиал"
        Case "мск"
            value = "Московский филиал"
        Case "гсг"
            value = "АО Гипроспецгаз"
        Case "ги"
            value = "ООО Газпром инвест"
        Case "мах"
            value = "Махачкалинский филиал"
        Case "под"
            value = "Подольский филиал"
        Case "сар"
            value = "Саратовский филиал"
        Case "тю"
            value = "Тюменский филиал"
        Case "нск"
            value = "Новосибирский филиал"
        Case "ст"
            value = "Ставропольский филиал"
        End Select
    ActiveCell.value = value
    val2 = value
    
    'От кого - ФИО
    Sheets("Вся входящая корреспонденция").Select
    Range("K4").Select
    Select Case val2
        Case "Нижегородский филиал"
            value = "Д.Г. Репин"
        Case "Санкт-Петербургский филиал"
            value = "Д.В. Яшков"
        Case "Московский филиал"
            value = "Н.В. Варламов"
        Case "АО Гипроспецгаз"
            value = "Е.А. Соловьев"
        Case "ООО Газпром инвест"
            value = "Л.И. Левченко"
        Case "Махачкалинский филиал"
            value = "И.Г. Гаджидадаев "
        Case "Подольский филиал"
            value = "А.В. Букин"
        Case "Саратовский филиал"
            value = "В.А. Вагарин"
        Case "Тюменский филиал"
            value = "С.А. Скрылев"
        Case "Новосибирский филиал"
            value = "А.А. Ковин"
        Case "Ставропольский филиал"
            value = "Р.Х. Гадиров"
        Case Else
            value = InputBox("ФИО От кого" & vbCrLf & "** = выход", "ОТ КОГО")
            If value = "**" Then Exit Sub
    End Select
    ActiveCell.value = value

    'Ответственный
    Sheets("Вся входящая корреспонденция").Select
    Range("N4").Select
    value = InputBox("Исполнитель" & vbCrLf & "1 = Т.А. Стаценко" & vbCrLf & "2 = Стаценко/Прокопьева" & vbCrLf & "3 = Т.В. Костина" & vbCrLf & "** = выход", "Исполнитель")
    Select Case value
        Case "**"
            Exit Sub
        Case 1
            value = "Т.А. Стаценко"
        Case 2
            value = "Т.А. Стаценко/ О.Г. Прокопьева"
        Case 3
            value = "Т.В. Костина"
        End Select
    ActiveCell.value = value
    
    
End Sub
