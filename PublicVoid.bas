Attribute VB_Name = "PublicVoid"
' Created by Gareev Ruslan, only for self using
' Version 2#3
Public varVendorNumber As String
Public Server As String, UserName As String, Password As String
Public wrkBging As Workspace
Public dbMaconomy As Database
Public rs As Recordset
Public dbName, dbPass As String
Public DoProcess As Boolean

Sub makeZBMmenu()
    Dim ZBMmenu As Boolean
    Dim NewItem As Object, rmenu As Object
    
    'add FTS menu if necessary
    BBMmenu = False
    For Each rmenu In MenuBars(xlWorksheet).Menus
        If rmenu.Caption = "Запросы в Maconomy" Then ZBMmenu = True
    Next rmenu
    
    If BBMmenu = False Then
        Set NewItem = MenuBars(xlWorksheet).Menus.Add(Caption:="Запросы в Maconomy")
        addZBMitems
    End If

End Sub
Sub makeBBMmenu()
    Dim BBMmenu As Boolean
    Dim NewItem As Object, rmenu As Object
    
    'add FTS menu if necessary
    BBMmenu = False
    For Each rmenu In MenuBars(xlWorksheet).Menus
        If rmenu.Caption = "Макросы Продукт Менеджера" Then BBMmenu = True
    Next rmenu
    
    If BBMmenu = False Then
        Set NewItem = MenuBars(xlWorksheet).Menus.Add(Caption:="Макросы Продукт Менеджера")
        addBBMitems
    End If

End Sub
Sub addZBMitems()
    Dim NewItem As Object
    With MenuBars(xlWorksheet).Menus("Запросы в Maconomy").MenuItems
        Set NewItem = .Add(Caption:="Запрос 'Замен'", OnAction:="UploadZam", Before:=5)
        
    End With
End Sub

Sub addBBMitems()
    Dim NewItem As Object
    With MenuBars(xlWorksheet).Menus("Макросы Продукт Менеджера").MenuItems
        Set NewItem = .Add(Caption:="Переименовать файлы", OnAction:="ConvertNewFileName", Before:=5)
        Set NewItem = .Add(Caption:="F2+Enter(Выдел.)", OnAction:="RefreshCellsFE", Before:=5)
        Set NewItem = .Add(Caption:="Лист скачивания картинок", OnAction:="CreateSheetForPictureDownload", Before:=5)
        Set NewItem = .Add(Caption:="Подсчет файлов в папке", OnAction:="FilenamesCollectionSb", Before:=5)
        Set NewItem = .Add(Caption:="Проверить Совпадения(Выдел.)", OnAction:="ПроверитьСовпадения", Before:=5)
        Set NewItem = .Add(Caption:="Разделить столбец по слову(Выдел.)", OnAction:="РазделениеПоСтолбцам", Before:=5)
        Set NewItem = .Add(Caption:="Скачать картинки с листа", OnAction:="Save_Object_As_Picture_NamesFromCells", Before:=5)
        Set NewItem = .Add(Caption:="Дублеры", OnAction:="bb", Before:=5)
        Set NewItem = .Add(Caption:="ГруппыДиапазоны", OnAction:="GroupStack", Before:=5)
        ' уберают одинаковые значения и оставляют 1но (первое).
    
    End With
End Sub

Sub ConvertNewFileName()
    If Len(ActiveSheet.[a2]) = 0 Then Exit Sub
    Dim V, Старое_имя$, Новое_имя$, Расширение$
    For Each V In range(ActiveSheet.[a2], ActiveSheet.[A1].End(xlDown)) 'у таблицы должен быть заголовок
        Старое_имя = V 'столбец "A"
        On Error Resume Next
        Расширение = ".jpg"
        Расширение = Mid(Старое_имя, InStrRev(Старое_имя, "."))
        Новое_имя = V(1, 2) 'стобец "B"
        'Подставляем старое расширение
        Новое_имя = Left(Новое_имя, InStrRev(Новое_имя, ".") - 1) & Расширение
        Err.Clear
        Name Старое_имя As Новое_имя
        If Err <> 0 Then Debug.Print "ERROR: " & Старое_имя & " ---> " & Новое_имя
        On Error GoTo 0
    Next
    MsgBox "Все файлы переименованы"
End Sub

Sub RefreshCellsFE()
Dim R As range, rr As range
Set rr = Selection
' Selection.NumberFormat = "@"
' If rr = 0 Then Debug.Print "ERROR: Выделите 1 столбец"
For Each R In rr
    Application.SendKeys "{F2}"
    Application.SendKeys "{ENTER}"
Next
End Sub


Sub CreateSheetForPictureDownload()
Sheets.Add
ActiveSheet.Name = "PictureDownload"
    ActiveSheet.Cells(1, 1).Value = "Название внутреннеый папки"
    Columns("A:A").ColumnWidth = 29.14
    ActiveSheet.Cells(1, 2).Value = "Наименование файла"
    Columns("B:B").ColumnWidth = 21.29
    ActiveSheet.Cells(1, 3).Value = "Ссылка"
    Columns("C:C").ColumnWidth = 39.57
    range("A1:C1").Select
    Selection.Font.Bold = True
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With
    
    ActiveSheet.Buttons.Add(500, 8.25, 123.75, 30.75).Select
    Selection.OnAction = "ОсновнойМакрос"
     ActiveSheet.Shapes.range(Array("Button 1")).Select
    Selection.Characters.Text = "Запуск"
    With Selection.Characters(Start:=1, Length:=6).Font
        .Name = "Calibri"
        .FontStyle = "обычный"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
End Sub

Sub FilenamesCollectionSb()
    ' Ищем на рабочем столе все файлы TXT, и выводим на лист список их имён.
    ' Просматриваются папки с глубиной вложения не более трёх.

    Dim coll As Collection, ПутьКПапке, Формат As String
    ' получаем путь к папке РАБОЧИЙ СТОЛ
    ПутьКПапке = InputBox("Пример : C:\Users\Vasya\Desktop\Google картинки", "Введите путь к папке")
    ' CreateObject("WScript.Shell").SpecialFolders ("Desktop")
    ' считываем в колекцию coll нужные имена файлов
    Формат = InputBox("Пример : .xls , .txt , .exe", "Введите формат искомого файла", ".jpg")
    Set coll = FilenamesCollection(ПутьКПапке, Формат, 1)

    Application.ScreenUpdating = False    ' отключаем обновление экрана
    ' создаём новую книгу
    Dim sh As Worksheet: Set sh = Workbooks.Add.Worksheets(1)
    ' формируем заголовки таблицы
    With sh.range("a1").Resize(, 3)
        .Value = Array("№", "Имя файла", "Полный путь")
        .Font.Bold = True: .Interior.ColorIndex = 17
    End With

    ' выводим результаты на лист
    For i = 1 To coll.Count ' перебираем все элементы коллекции, содержащей пути к файлам
        sh.range("a" & sh.Rows.Count).End(xlUp).Offset(1).Resize(, 3).Value = _
        Array(i, Dir(coll(i)), coll(i))    ' выводим на лист очередную строку
        DoEvents    ' временно передаём управление ОС
    Next
    sh.range("a:c").EntireColumn.AutoFit    ' автоподбор ширины столбцов
    [a2].Activate: ActiveWindow.FreezePanes = True ' закрепляем первую строку листа
End Sub

' Сюда вствлять новый Саб, для высплыв. Окна.

Function СцепитьМного(Диапазон As range, Optional Разделитель As String = " ", Optional БезПовторов As Boolean = False)
    Dim avData, lr As Long, lc As Long, sRes As String
    avData = Диапазон.Value
    If Not IsArray(avData) Then
        СцепитьМного = avData
        Exit Function
    End If
 
    For lc = 1 To UBound(avData, 2)
        For lr = 1 To UBound(avData, 1)
            If Len(avData(lr, lc)) Then
                sRes = sRes & Разделитель & avData(lr, lc)
            End If
        Next lr
    Next lc
    If Len(sRes) Then
        sRes = Mid(sRes, Len(Разделитель) + 1)
    End If
    
    If БезПовторов Then
        Dim oDict As Object, sTmpStr
        Set oDict = CreateObject("Scripting.Dictionary")
        sTmpStr = Split(sRes, Разделитель)
        On Error Resume Next
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            oDict.Add sTmpStr(lr), sTmpStr(lr)
        Next lr
        sRes = ""
        sTmpStr = oDict.keys
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            sRes = sRes & IIf(sRes <> "", Разделитель, "") & sTmpStr(lr)
        Next lr
    End If
    СцепитьМного = sRes
End Function

Function ЦветПодтверждения(ЯчейкаПроверки1 As range, ЯчейкаПроверки2 As range)
Application.Volatile True
If ЯчейкаПроверки1.Interior.Color = 65535 And ЯчейкаПроверки2.Interior.Color = 65535 And ЯчейкаПроверки1 <> "-" And ЯчейкаПроверки2 <> "-" Then
ЦветПодтверждения = "Подтвержден/Подтвержден"
Else
    If ЯчейкаПроверки1.Interior.Color = 65535 And ЯчейкаПроверки1 <> "-" Then
        If ЯчейкаПроверки2 = "-" Then
            ЦветПодтверждения = "Подтвержден/-"
        Else
            ЦветПодтверждения = "Подтвержден/Не подтвержден"
        End If
    Else
        If ЯчейкаПроверки2.Interior.Color = 65535 And ЯчейкаПроверки2 <> "-" Then
            If ЯчейкаПроверки1 = "-" Then
                ЦветПодтверждения = "-/Подтвержден"
            Else
                ЦветПодтверждения = "Не подтвержден/Подтвержден"
            End If
        Else
            If ЯчейкаПроверки1 = "-" Then
                ЦветПодтверждения = "-/Не подтвержден"
            Else
                If ЯчейкаПроверки2 = "-" Then
                    ЦветПодтверждения = "Не подтвержден/-"
                Else
                    ЦветПодтверждения = "Не подтвержден/Не подтвержден"
                End If
            End If
        
        End If
    End If
End If

End Function

Public Function ИзвлечьГиперссылку(ByVal range As range) As String
 If (range.Hyperlinks.Count > 0) Then
 ИзвлечьГиперссылку = range.Hyperlinks(1).Address
 Else
 ИзвлечьГиперссылку = ""
 End If
 End Function

Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                             Optional ByVal SearchDeep As Long = 999) As Collection
    ' Получает в качестве параметра путь к папке FolderPath,
    ' маску имени искомых файлов Mask (будут отобраны только файлы с такой маской/расширением)
    ' и глубину поиска SearchDeep в подпапках (если SearchDeep=1, то подпапки не просматриваются).
    ' Возвращает коллекцию, содержащую полные пути найденных файлов
    ' (применяется рекурсивный вызов процедуры GetAllFileNamesUsingFSO)

    Set FilenamesCollection = New Collection    ' создаём пустую коллекцию
    Set FSO = CreateObject("Scripting.FileSystemObject")    ' создаём экземпляр FileSystemObject
    GetAllFileNamesUsingFSO FolderPath, Mask, FSO, FilenamesCollection, SearchDeep ' поиск
    Set FSO = Nothing: Application.StatusBar = False    ' очистка строки состояния Excel
End Function

Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                                 ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
    ' перебирает все файлы и подпапки в папке FolderPath, используя объект FSO
    ' перебор папок осуществляется в том случае, если SearchDeep > 1
    ' добавляет пути найденных файлов в коллекцию FileNamesColl
    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
    If Not curfold Is Nothing Then    ' если удалось получить доступ к папке

        ' раскомментируйте эту строку для вывода пути к просматриваемой
        ' в текущий момент папке в строку состояния Excel
        Application.StatusBar = "Поиск в папке: " & FolderPath

        For Each fil In curfold.Files    ' перебираем все файлы в папке FolderPath
            If fil.Name Like "*" & Mask Then FileNamesColl.Add fil.Path
        Next
        SearchDeep = SearchDeep - 1    ' уменьшаем глубину поиска в подпапках
        If SearchDeep Then    ' если надо искать глубже
            For Each sfol In curfold.SubFolders    ' ' перебираем все подпапки в папке FolderPath
                GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
            Next
        End If
        Set fil = Nothing: Set curfold = Nothing    ' очищаем переменные
    End If
End Function

Sub ПроверитьСовпадения()
Dim SearchString, CCC, NewCreation
Const ForReading = 1
Dim objFSO, objTextFile, a, b, c, i&
    c = 1
    Dim pi As New ProgressIndicator
    pi.Show "Поиск подходящих синонимов"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("Z:\Market\ПМ\ПМ\AnSwitch.txt", ForReading)
a = Split(objTextFile.ReadAll, vbNewLine)
objTextFile.Close
Set objTextFile = Nothing
Set SearchString = Selection
NewCreation = True
pi.StartNewAction 0, 100, "Поиск совпадений", , , SearchString.Count
For Each CCC In SearchString
    If CCC.Cells.Value <> 0 Then
    If CCC.Cells(1, 2).Value <> 0 Then
    Else
    If NewCreation Then
    ReDim b(0 To UBound(a), 1 To 2)
    End If
        For i = 0 To UBound(a) - 1
            If NewCreation Then
                b(i, 1) = Split(a(i), "*")(0)
                b(i, 2) = Split(a(i), "*")(1)
            End If
            Pos = InStr(CCC.Cells.Value, b(i, 1))
            If (Pos > 0) Then
                If CCC.Cells(1, 2).Value = 0 Then
                    CCC.Cells(1, 2).Value = b(i, 2)
                    Else
                    CCC.Cells(1, 2).Value = CCC.Cells(1, 2).Value & "; " & b(i, 2)
                End If
            End If
            If i = UBound(a) - 1 And CCC.Cells(1, 2).Value = 0 Then
                CCC.Cells(1, 2).Value = "Нет Совпадений"
            End If
        Next
    End If
    End If
    pi.SubAction , "Обрабатывается ячейка " & c & " из " & SearchString.Count
    c = c + 1
NewCreation = False
Next
pi.Hide
End Sub

Sub РазделениеПоСтолбцам()
Dim SearchString, CCC, NewCreation
Dim СловоПоиска
Const ForReading = 1
Dim x, y, z, c
    c = 1
'Dim pi As New ProgressIndicator
'pi.Show "Поиск подходящих синонимов"

СловоПоиска = InputBox("Введите слово для поиска", " ")
x = Len(СловоПоиска)

Set SearchString = Selection
'pi.StartNewAction 0, 100, "Поиск совпадений", , , SearchString.Count

For Each CCC In SearchString
    If CCC.Cells.Value <> 0 Then
    If CCC.Cells(1, 2).Value <> 0 Then
    Else
    y = Len(CCC.Cells.Value)
    Pos = InStr(CCC.Cells.Value, СловоПоиска)
    If Pos <> 0 Then
    z = Left(CCC.Cells.Value, Pos + x)
    CCC.Cells(1, 2).Value = Right(CCC.Cells.Value, y - (Pos + x))
    CCC.Cells.Value = z
    Else
    End If
    End If
    End If
    'pi.SubAction , "Обрабатывается ячейка " & c & " из " & SearchString.Count
    'c = c + 1

Next
'pi.Hide
End Sub


Sub Save_Object_As_Picture_NamesFromCells()
    Dim li As Long, oObj As Shape, wsSh As Worksheet, wsTmpSh As Worksheet
    Dim sImagesPath As String, sName As String
    Dim lNamesCol As Long, s As String
    
    s = InputBox("Укажите номер столбца с именами для картинок" & vbNewLine & _
                 "(0 - столбец в котором сама картинка)", "www.excel-vba.ru", "")
    If StrPtr(s) = 0 Then Exit Sub
    lNamesCol = Val(s)
    
    sImagesPath = ActiveWorkbook.Path & "\images\" '"
    If Dir(sImagesPath, 16) = "" Then
        MkDir sImagesPath
    End If
'    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Set wsSh = ActiveSheet
    Set wsTmpSh = ActiveWorkbook.Sheets.Add
    For Each oObj In wsSh.Shapes
        If oObj.Type = 13 Then
            oObj.Copy
            If lNamesCol = 0 Then
                sName = oObj.TopLeftCell.Value
            Else
                sName = wsSh.Cells(oObj.TopLeftCell.row, lNamesCol).Value
            End If
            'если в ячейке были символы, запрещенные
            'для использования в качестве имен для файлов - удаляем
            sName = CheckName(sName)
            'если sName в результате пусто - даем имя unnamed_ с порядковым номером
            If sName = "" Then
                li = li + 1
                sName = "unnamed_" & li
            End If
            With wsTmpSh.ChartObjects.Add(0, 0, oObj.Width, oObj.Height).Chart
                .ChartArea.Border.LineStyle = 0
                .Paste
                .Export Filename:=sImagesPath & sName & ".jpg", FilterName:="JPG"
                .Parent.Delete
            End With
        End If
    Next oObj
    Set oObj = Nothing: Set wsSh = Nothing
    wsTmpSh.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Объекты сохранены в папке: " & sImagesPath, vbInformation, "www.excel-vba.ru"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CheckName
' Purpose   : Функция проверки правильности имени
'---------------------------------------------------------------------------------------
Function CheckName(sName As String)
    Dim objRegExp As Object
    Dim s As String
    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Global = True: objRegExp.IgnoreCase = True
    objRegExp.Pattern = "[:,\\,/,?,\*,\<,\>,\',\|,""""]"
    s = objRegExp.Replace(sName, "")
    CheckName = s
End Function

Sub GroupStack()
Dim СтолбецГрупп, Cre, LastGroup, i As Long
Dim МРЦ, СК, Art, GroupSel
Dim Rng As range

'Set СтолбецГрупп = Selection
'Mycount = Selection.Count

'Set Rng = range(Cells(Art.Cells.row + i, 4), GroupSel.Cells(1, 4))

On Error Resume Next
Set Art = Application.InputBox _
(prompt:="Выбрать Артикул:", Type:=8)
If Art Is Nothing Then Exit Sub

On Error Resume Next
Set МРЦ = Application.InputBox _
(prompt:="Выбрать РОЦ1:", Type:=8)
If МРЦ Is Nothing Then Exit Sub

On Error Resume Next
Set СК = Application.InputBox _
(prompt:="Выбрать Скидку клиента:", Type:=8)
If СК Is Nothing Then Exit Sub

Art.Cells(1, 2).EntireColumn.Insert
Art.Cells(1, 2).EntireColumn.Insert

i = 2
Set GroupSel = Cells(Art.Cells.row + i, 1)
Do While GroupSel <> 0

    If GroupSel.Cells(1, 2).Value = 0 Then
        LastGroup = GroupSel.Value
    Else
        GroupSel.Cells(1, 3).Value = LastGroup & " " & GroupSel.Cells(1, МРЦ.Cells.Column).Value & " " & GroupSel.Cells(1, СК.Cells.Column).Value
    End If
    i = i + 1
    Set GroupSel = Cells(Art.Cells.row + i, 1)
    
Loop

Cells.Select
With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.UnMerge

range(Cells(Art.Cells.row + 3, Art.Cells.Column + 1), GroupSel.Cells(1, Art.Cells.Column + 1)).Select
ПроверитьСовпадения
End Sub



'ZBM Макросы:
Sub UploadZam()
Dim sql As String, item As String
Dim R As range

Dim connectStr As String
Dim varItemNumberORA As String
Dim j As Integer
Dim i As Integer
Dim n As Integer
Dim row As Integer
Dim theEnd As Boolean
    ConnectDB

'    On Error GoTo DB_Error
row = 3
             
ActiveWindow.RangeSelection.Rows.Cells(1, 2).EntireColumn.Insert
ActiveWindow.RangeSelection.Rows.Cells(1, 2).EntireColumn.Insert
ActiveWindow.RangeSelection.Rows.Cells(1, 2).EntireColumn.Insert
ActiveWindow.RangeSelection.Rows.Cells(1, 2).EntireColumn.Insert
    
    For Each R In ActiveWindow.RangeSelection.Rows
        If R.Cells(1, 1).Value <> "" Then
            item = CStr(R.Cells(1, 1).Value)
            j = 1
            varItemNumberORA = ""
            Do While j <= Len(item) And j <= 255
                varItemNumberORA = varItemNumberORA + "CHR (" + Trim(Str(Asc(Mid(item, j, 1)))) + ")"
                If j < Len(item) Then
                    varItemNumberORA = varItemNumberORA + " || "
                End If
                j = j + 1
            Loop
            If item = "" Then
                varItemNumberORA = "CHR(32)"
            End If
            i = 2
           sql = "select distinct SubstitutionItemNumber, NVL(ctype.name,' ') ITEMTYPE, decode(i.Blocked, 1, 'yes', 'no') Blocked, i.SupplementaryText1 " _
                + "from " + dbName + "ItemInformation i, " + dbName + "PopupItem CType " _
                + "where ITEMNUMBER=" + varItemNumberORA _
                + "and CType.PopupItemNumber(+)=i.ITEMType and (CType.PopupTypeName='ItemTypeType' or CType.PopupTypeName is null) order by itemnumber"
                
            Set rs = dbMaconomy.OpenRecordset(sql, dbReadOnly)
  
            Do While Not rs.EOF
                R.Cells(1, 2).Value = rs!ITEMTYPE
                R.Cells(1, 3).Value = rs!Blocked
                R.Cells(1, 4).Value = rs!SupplementaryText1
                R.Cells(1, 5).Value = rs!SubstitutionItemNumber
                rs.MoveNext
            Loop
            
            theEnd = False
            i = 5
            Do While theEnd = False
                If R.Cells(1, i).Value <> "" And R.Cells(1, 1).Value <> R.Cells(1, i).Value Then
                    item = CStr(R.Cells(1, i).Value)
                    j = 1
                
                    If row < i Then
                    row = i
                    ActiveWindow.RangeSelection.Rows.Cells(1, i + 1).EntireColumn.Insert
                    End If
                
                    
                
                    varItemNumberORA = ""
                    Do While j <= Len(item) And j <= 255
                        varItemNumberORA = varItemNumberORA + "CHR (" + Trim(Str(Asc(Mid(item, j, 1)))) + ")"
                        If j < Len(item) Then
                            varItemNumberORA = varItemNumberORA + " || "
                        End If
                        j = j + 1
                    Loop
                    If item = "" Then
                        varItemNumberORA = "CHR(32)"
                    End If
            
                        sql = "select SubstitutionItemNumber " _
                        + "from " + dbName + "ItemInformation " _
                        + "where ITEMNUMBER=" + varItemNumberORA
                
                    Set rs = dbMaconomy.OpenRecordset(sql, dbReadOnly)
  
                    Do While Not rs.EOF
                        R.Cells(1, i + 1).Value = rs!SubstitutionItemNumber
                        rs.MoveNext
                    Loop
                    
                    For n = 5 To i - 1
                        If R.Cells(1, n).Value = R.Cells(1, i + 1).Value Then
                            theEnd = True
                            R.Cells(1, n).Interior.Color = 65535
                            R.Cells(1, i + 1).Interior.Color = 65535
                        End If
                    Next n
                    
                    i = i + 1
                
                    If i = 15 Then
                        theEnd = True
                        R.Cells(1, i + 1).Value = "Больше 15 повторений"
                    End If
                
                Else
                    theEnd = True
                    R.Cells(1, 1).Interior.Color = 65535
                    R.Cells(1, i).Interior.Color = 65535
                End If
            Loop
            
        End If
    Next R
    DisconnectDB
'Generic Error Handler
DB_Error:
'   Display the error message and then exit
    If (Err.Number <> 0) Then
        MsgBox Err.Description
    End If
End Sub

Sub bb()
Dim c As range, x
With CreateObject("scripting.dictionary")
  For Each c In Selection
    .RemoveAll
    For Each x In Split(c)
      .item(x) = 0
    Next
    c = Join(.keys)
  Next
End With
End Sub

Sub CombineWorkbooks()
    Dim FilesToOpen
    Dim x As Integer
 
    Application.ScreenUpdating = False  'отключаем обновление экрана для скорости
     
    'вызываем диалог выбора файлов для импорта
    FilesToOpen = Application.GetOpenFilename _
      (FileFilter:="All files (*.*), *.*", _
      MultiSelect:=True, Title:="Files to Merge")
 
    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "Не выбрано ни одного файла!"
        Exit Sub
    End If
     
    'проходим по всем выбранным файлам
    x = 1
    While x <= UBound(FilesToOpen)
        Set importWB = Workbooks.Open(Filename:=FilesToOpen(x))
        Sheets().Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        importWB.Close savechanges:=False
        x = x + 1
    Wend
 
    Application.ScreenUpdating = True
End Sub

Sub Собрать_csv_файлы()
    ' Макрос копирует все csv файлы из заданной папки в один
    Dim TextLine$, MyPath$, MyFileName$, Usl As Boolean
    MyPath = "C:\Users\"
    MyFileName = Dir(MyPath & "*.csv")
    'MkDir "C:\Users\"
    Open MyPath & "свод\свод1.csv" For Output As #1
    Usl = True
    Do Until MyFileName = ""
        Open MyPath & MyFileName For Input Lock Read As #2
        Line Input #2, TextLine
        If Usl Then Print #1, TextLine: Usl = False
        Do While Not EOF(2)
            Line Input #2, TextLine
            Print #1, TextLine
        Loop
        Close #2
        MyFileName = Dir
    Loop
    Close #1
End Sub
