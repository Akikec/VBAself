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
        If rmenu.Caption = "������� � Maconomy" Then ZBMmenu = True
    Next rmenu
    
    If BBMmenu = False Then
        Set NewItem = MenuBars(xlWorksheet).Menus.Add(Caption:="������� � Maconomy")
        addZBMitems
    End If

End Sub
Sub makeBBMmenu()
    Dim BBMmenu As Boolean
    Dim NewItem As Object, rmenu As Object
    
    'add FTS menu if necessary
    BBMmenu = False
    For Each rmenu In MenuBars(xlWorksheet).Menus
        If rmenu.Caption = "������� ������� ���������" Then BBMmenu = True
    Next rmenu
    
    If BBMmenu = False Then
        Set NewItem = MenuBars(xlWorksheet).Menus.Add(Caption:="������� ������� ���������")
        addBBMitems
    End If

End Sub
Sub addZBMitems()
    Dim NewItem As Object
    With MenuBars(xlWorksheet).Menus("������� � Maconomy").MenuItems
        Set NewItem = .Add(Caption:="������ '�����'", OnAction:="UploadZam", Before:=5)
        
    End With
End Sub

Sub addBBMitems()
    Dim NewItem As Object
    With MenuBars(xlWorksheet).Menus("������� ������� ���������").MenuItems
        Set NewItem = .Add(Caption:="������������� �����", OnAction:="ConvertNewFileName", Before:=5)
        Set NewItem = .Add(Caption:="F2+Enter(�����.)", OnAction:="RefreshCellsFE", Before:=5)
        Set NewItem = .Add(Caption:="���� ���������� ��������", OnAction:="CreateSheetForPictureDownload", Before:=5)
        Set NewItem = .Add(Caption:="������� ������ � �����", OnAction:="FilenamesCollectionSb", Before:=5)
        Set NewItem = .Add(Caption:="��������� ����������(�����.)", OnAction:="�������������������", Before:=5)
        Set NewItem = .Add(Caption:="��������� ������� �� �����(�����.)", OnAction:="��������������������", Before:=5)
        Set NewItem = .Add(Caption:="������� �������� � �����", OnAction:="Save_Object_As_Picture_NamesFromCells", Before:=5)
        Set NewItem = .Add(Caption:="�������", OnAction:="bb", Before:=5)
        Set NewItem = .Add(Caption:="���������������", OnAction:="GroupStack", Before:=5)
        ' ������� ���������� �������� � ��������� 1�� (������).
    
    End With
End Sub

Sub ConvertNewFileName()
    If Len(ActiveSheet.[a2]) = 0 Then Exit Sub
    Dim V, ������_���$, �����_���$, ����������$
    For Each V In range(ActiveSheet.[a2], ActiveSheet.[A1].End(xlDown)) '� ������� ������ ���� ���������
        ������_��� = V '������� "A"
        On Error Resume Next
        ���������� = ".jpg"
        ���������� = Mid(������_���, InStrRev(������_���, "."))
        �����_��� = V(1, 2) '������ "B"
        '����������� ������ ����������
        �����_��� = Left(�����_���, InStrRev(�����_���, ".") - 1) & ����������
        Err.Clear
        Name ������_��� As �����_���
        If Err <> 0 Then Debug.Print "ERROR: " & ������_��� & " ---> " & �����_���
        On Error GoTo 0
    Next
    MsgBox "��� ����� �������������"
End Sub

Sub RefreshCellsFE()
Dim R As range, rr As range
Set rr = Selection
' Selection.NumberFormat = "@"
' If rr = 0 Then Debug.Print "ERROR: �������� 1 �������"
For Each R In rr
    Application.SendKeys "{F2}"
    Application.SendKeys "{ENTER}"
Next
End Sub


Sub CreateSheetForPictureDownload()
Sheets.Add
ActiveSheet.Name = "PictureDownload"
    ActiveSheet.Cells(1, 1).Value = "�������� ����������� �����"
    Columns("A:A").ColumnWidth = 29.14
    ActiveSheet.Cells(1, 2).Value = "������������ �����"
    Columns("B:B").ColumnWidth = 21.29
    ActiveSheet.Cells(1, 3).Value = "������"
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
    Selection.OnAction = "��������������"
     ActiveSheet.Shapes.range(Array("Button 1")).Select
    Selection.Characters.Text = "������"
    With Selection.Characters(Start:=1, Length:=6).Font
        .Name = "Calibri"
        .FontStyle = "�������"
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
    ' ���� �� ������� ����� ��� ����� TXT, � ������� �� ���� ������ �� ���.
    ' ��������������� ����� � �������� �������� �� ����� ���.

    Dim coll As Collection, ����������, ������ As String
    ' �������� ���� � ����� ������� ����
    ���������� = InputBox("������ : C:\Users\Vasya\Desktop\Google ��������", "������� ���� � �����")
    ' CreateObject("WScript.Shell").SpecialFolders ("Desktop")
    ' ��������� � �������� coll ������ ����� ������
    ������ = InputBox("������ : .xls , .txt , .exe", "������� ������ �������� �����", ".jpg")
    Set coll = FilenamesCollection(����������, ������, 1)

    Application.ScreenUpdating = False    ' ��������� ���������� ������
    ' ������ ����� �����
    Dim sh As Worksheet: Set sh = Workbooks.Add.Worksheets(1)
    ' ��������� ��������� �������
    With sh.range("a1").Resize(, 3)
        .Value = Array("�", "��� �����", "������ ����")
        .Font.Bold = True: .Interior.ColorIndex = 17
    End With

    ' ������� ���������� �� ����
    For i = 1 To coll.Count ' ���������� ��� �������� ���������, ���������� ���� � ������
        sh.range("a" & sh.Rows.Count).End(xlUp).Offset(1).Resize(, 3).Value = _
        Array(i, Dir(coll(i)), coll(i))    ' ������� �� ���� ��������� ������
        DoEvents    ' �������� ������� ���������� ��
    Next
    sh.range("a:c").EntireColumn.AutoFit    ' ���������� ������ ��������
    [a2].Activate: ActiveWindow.FreezePanes = True ' ���������� ������ ������ �����
End Sub

' ���� �������� ����� ���, ��� �������. ����.

Function ������������(�������� As range, Optional ����������� As String = " ", Optional ����������� As Boolean = False)
    Dim avData, lr As Long, lc As Long, sRes As String
    avData = ��������.Value
    If Not IsArray(avData) Then
        ������������ = avData
        Exit Function
    End If
 
    For lc = 1 To UBound(avData, 2)
        For lr = 1 To UBound(avData, 1)
            If Len(avData(lr, lc)) Then
                sRes = sRes & ����������� & avData(lr, lc)
            End If
        Next lr
    Next lc
    If Len(sRes) Then
        sRes = Mid(sRes, Len(�����������) + 1)
    End If
    
    If ����������� Then
        Dim oDict As Object, sTmpStr
        Set oDict = CreateObject("Scripting.Dictionary")
        sTmpStr = Split(sRes, �����������)
        On Error Resume Next
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            oDict.Add sTmpStr(lr), sTmpStr(lr)
        Next lr
        sRes = ""
        sTmpStr = oDict.keys
        For lr = LBound(sTmpStr) To UBound(sTmpStr)
            sRes = sRes & IIf(sRes <> "", �����������, "") & sTmpStr(lr)
        Next lr
    End If
    ������������ = sRes
End Function

Function �����������������(��������������1 As range, ��������������2 As range)
Application.Volatile True
If ��������������1.Interior.Color = 65535 And ��������������2.Interior.Color = 65535 And ��������������1 <> "-" And ��������������2 <> "-" Then
����������������� = "�����������/�����������"
Else
    If ��������������1.Interior.Color = 65535 And ��������������1 <> "-" Then
        If ��������������2 = "-" Then
            ����������������� = "�����������/-"
        Else
            ����������������� = "�����������/�� �����������"
        End If
    Else
        If ��������������2.Interior.Color = 65535 And ��������������2 <> "-" Then
            If ��������������1 = "-" Then
                ����������������� = "-/�����������"
            Else
                ����������������� = "�� �����������/�����������"
            End If
        Else
            If ��������������1 = "-" Then
                ����������������� = "-/�� �����������"
            Else
                If ��������������2 = "-" Then
                    ����������������� = "�� �����������/-"
                Else
                    ����������������� = "�� �����������/�� �����������"
                End If
            End If
        
        End If
    End If
End If

End Function

Public Function ������������������(ByVal range As range) As String
 If (range.Hyperlinks.Count > 0) Then
 ������������������ = range.Hyperlinks(1).Address
 Else
 ������������������ = ""
 End If
 End Function

Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                             Optional ByVal SearchDeep As Long = 999) As Collection
    ' �������� � �������� ��������� ���� � ����� FolderPath,
    ' ����� ����� ������� ������ Mask (����� �������� ������ ����� � ����� ������/�����������)
    ' � ������� ������ SearchDeep � ��������� (���� SearchDeep=1, �� �������� �� ���������������).
    ' ���������� ���������, ���������� ������ ���� ��������� ������
    ' (����������� ����������� ����� ��������� GetAllFileNamesUsingFSO)

    Set FilenamesCollection = New Collection    ' ������ ������ ���������
    Set FSO = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject
    GetAllFileNamesUsingFSO FolderPath, Mask, FSO, FilenamesCollection, SearchDeep ' �����
    Set FSO = Nothing: Application.StatusBar = False    ' ������� ������ ��������� Excel
End Function

Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                                 ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
    ' ���������� ��� ����� � �������� � ����� FolderPath, ��������� ������ FSO
    ' ������� ����� �������������� � ��� ������, ���� SearchDeep > 1
    ' ��������� ���� ��������� ������ � ��������� FileNamesColl
    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
    If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����

        ' ���������������� ��� ������ ��� ������ ���� � ���������������
        ' � ������� ������ ����� � ������ ��������� Excel
        Application.StatusBar = "����� � �����: " & FolderPath

        For Each fil In curfold.Files    ' ���������� ��� ����� � ����� FolderPath
            If fil.Name Like "*" & Mask Then FileNamesColl.Add fil.Path
        Next
        SearchDeep = SearchDeep - 1    ' ��������� ������� ������ � ���������
        If SearchDeep Then    ' ���� ���� ������ ������
            For Each sfol In curfold.SubFolders    ' ' ���������� ��� �������� � ����� FolderPath
                GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
            Next
        End If
        Set fil = Nothing: Set curfold = Nothing    ' ������� ����������
    End If
End Function

Sub �������������������()
Dim SearchString, CCC, NewCreation
Const ForReading = 1
Dim objFSO, objTextFile, a, b, c, i&
    c = 1
    Dim pi As New ProgressIndicator
    pi.Show "����� ���������� ���������"
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile("Z:\Market\��\��\AnSwitch.txt", ForReading)
a = Split(objTextFile.ReadAll, vbNewLine)
objTextFile.Close
Set objTextFile = Nothing
Set SearchString = Selection
NewCreation = True
pi.StartNewAction 0, 100, "����� ����������", , , SearchString.Count
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
                CCC.Cells(1, 2).Value = "��� ����������"
            End If
        Next
    End If
    End If
    pi.SubAction , "�������������� ������ " & c & " �� " & SearchString.Count
    c = c + 1
NewCreation = False
Next
pi.Hide
End Sub

Sub ��������������������()
Dim SearchString, CCC, NewCreation
Dim �����������
Const ForReading = 1
Dim x, y, z, c
    c = 1
'Dim pi As New ProgressIndicator
'pi.Show "����� ���������� ���������"

����������� = InputBox("������� ����� ��� ������", " ")
x = Len(�����������)

Set SearchString = Selection
'pi.StartNewAction 0, 100, "����� ����������", , , SearchString.Count

For Each CCC In SearchString
    If CCC.Cells.Value <> 0 Then
    If CCC.Cells(1, 2).Value <> 0 Then
    Else
    y = Len(CCC.Cells.Value)
    Pos = InStr(CCC.Cells.Value, �����������)
    If Pos <> 0 Then
    z = Left(CCC.Cells.Value, Pos + x)
    CCC.Cells(1, 2).Value = Right(CCC.Cells.Value, y - (Pos + x))
    CCC.Cells.Value = z
    Else
    End If
    End If
    End If
    'pi.SubAction , "�������������� ������ " & c & " �� " & SearchString.Count
    'c = c + 1

Next
'pi.Hide
End Sub


Sub Save_Object_As_Picture_NamesFromCells()
    Dim li As Long, oObj As Shape, wsSh As Worksheet, wsTmpSh As Worksheet
    Dim sImagesPath As String, sName As String
    Dim lNamesCol As Long, s As String
    
    s = InputBox("������� ����� ������� � ������� ��� ��������" & vbNewLine & _
                 "(0 - ������� � ������� ���� ��������)", "www.excel-vba.ru", "")
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
            '���� � ������ ���� �������, �����������
            '��� ������������� � �������� ���� ��� ������ - �������
            sName = CheckName(sName)
            '���� sName � ���������� ����� - ���� ��� unnamed_ � ���������� �������
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
    MsgBox "������� ��������� � �����: " & sImagesPath, vbInformation, "www.excel-vba.ru"
End Sub
'---------------------------------------------------------------------------------------
' Procedure : CheckName
' Purpose   : ������� �������� ������������ �����
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
Dim ������������, Cre, LastGroup, i As Long
Dim ���, ��, Art, GroupSel
Dim Rng As range

'Set ������������ = Selection
'Mycount = Selection.Count

'Set Rng = range(Cells(Art.Cells.row + i, 4), GroupSel.Cells(1, 4))

On Error Resume Next
Set Art = Application.InputBox _
(prompt:="������� �������:", Type:=8)
If Art Is Nothing Then Exit Sub

On Error Resume Next
Set ��� = Application.InputBox _
(prompt:="������� ���1:", Type:=8)
If ��� Is Nothing Then Exit Sub

On Error Resume Next
Set �� = Application.InputBox _
(prompt:="������� ������ �������:", Type:=8)
If �� Is Nothing Then Exit Sub

Art.Cells(1, 2).EntireColumn.Insert
Art.Cells(1, 2).EntireColumn.Insert

i = 2
Set GroupSel = Cells(Art.Cells.row + i, 1)
Do While GroupSel <> 0

    If GroupSel.Cells(1, 2).Value = 0 Then
        LastGroup = GroupSel.Value
    Else
        GroupSel.Cells(1, 3).Value = LastGroup & " " & GroupSel.Cells(1, ���.Cells.Column).Value & " " & GroupSel.Cells(1, ��.Cells.Column).Value
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
�������������������
End Sub



'ZBM �������:
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
                        R.Cells(1, i + 1).Value = "������ 15 ����������"
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
 
    Application.ScreenUpdating = False  '��������� ���������� ������ ��� ��������
     
    '�������� ������ ������ ������ ��� �������
    FilesToOpen = Application.GetOpenFilename _
      (FileFilter:="All files (*.*), *.*", _
      MultiSelect:=True, Title:="Files to Merge")
 
    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "�� ������� �� ������ �����!"
        Exit Sub
    End If
     
    '�������� �� ���� ��������� ������
    x = 1
    While x <= UBound(FilesToOpen)
        Set importWB = Workbooks.Open(Filename:=FilesToOpen(x))
        Sheets().Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
        importWB.Close savechanges:=False
        x = x + 1
    Wend
 
    Application.ScreenUpdating = True
End Sub

Sub �������_csv_�����()
    ' ������ �������� ��� csv ����� �� �������� ����� � ����
    Dim TextLine$, MyPath$, MyFileName$, Usl As Boolean
    MyPath = "C:\Users\"
    MyFileName = Dir(MyPath & "*.csv")
    'MkDir "C:\Users\"
    Open MyPath & "����\����1.csv" For Output As #1
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
