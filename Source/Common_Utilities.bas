Attribute VB_Name = "Common_Utilities"
'''''==============================================================================================================================================
'''''                                            Platform-independent Functions
'''''==============================================================================================================================================
Sub AddMenuItem(itemText As String, itemCommand As String, itemFaceId As Long)
    ' Check if we have already added the menu item
    Dim initialized As Boolean
    Dim bef As Integer
    initialized = False
    bef = 1
    Dim Menu As CommandBars
    Set Menu = Application.CommandBars
    For i = 1 To Menu("Insert").Controls.Count
        With Menu("Insert").Controls(i)
            If .Caption = itemText Then
                initialized = True
                Exit For
            ElseIf InStr(.Caption, "Dia&gram") Then
                bef = i
            End If
        End With
    Next
    
    ' Create the menu choice.
    If Not initialized Then
        Dim NewControl As CommandBarControl
        Set NewControl = Menu("Insert").Controls.Add _
                              (Type:=msoControlButton, _
                               before:=bef, _
                               id:=itemFaceId)
        NewControl.Caption = itemText
        NewControl.OnAction = itemCommand
        NewControl.Style = msoButton
    End If
End Sub


Sub RemoveMenuItem(itemText As String)
    Dim Menu As CommandBars
    Set Menu = Application.CommandBars
    For i = 1 To Menu("Insert").Controls.Count
        If Menu("Insert").Controls(i).Caption = itemText Then
            Menu("Insert").Controls(i).Delete
            Exit For
        End If
    Next
    

End Sub

Sub ReadLaTeXFromFile(code As String, TempDir As String, FilePrefix As String)
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    
    'UseUTF8
    Dim objStream, strData

    Set objStream = CreateObject("ADODB.Stream")
    
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (TempDir & FilePrefix & ".tex")
    
    code = objStream.ReadText()
    
    objStream.Close
    Set objStream = Nothing
End Sub

Sub WriteLaTeXToFile(code As String, TempDir As String, FilePrefix As String)
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    
    'UseUTF8
    Dim BinaryStream As Object
    Set BinaryStream = CreateObject("ADODB.stream")
    BinaryStream.Type = 1
    BinaryStream.Open
    Dim adodbStream  As Object
    Set adodbStream = CreateObject("ADODB.Stream")
    With adodbStream
        .Type = 2 'Stream type
        .Charset = "utf-8"
        .Open
        .WriteText code
        '.SaveToFile TempDir & FilePrefix & ".tex", 2 'Save binary data To disk; problem: this includes a BOM
        ' Workaround to avoid BOM in file:
        .Position = 3 'skip BOM
        .CopyTo BinaryStream
        .Flush
        .Close
    End With
    BinaryStream.SaveToFile TempDir & FilePrefix & ".tex", 2 'Save binary data To disk
    BinaryStream.Flush
    BinaryStream.Close
    Set fs = Nothing
End Sub

Function QuotedStr(str As String)
    QuotedStr = """" & str & """"
End Function

Public Function EndWith(str As String, ending As String) As Boolean
     Dim endingLen As Integer
     endingLen = Len(ending)
     EndWith = (Right(Trim(str), endingLen) = ending)
End Function

Public Function StartWith(str As String, start As String) As Boolean
     Dim startLen As Integer
     startLen = Len(start)
     StartWith = (Left(Trim(str), startLen) = start)
End Function

Function IsInArray(arr As Variant, valueToCheck As String) As Boolean
    IsInArray = False
    For Each n In arr
        If n = valueToCheck Then
            IsInArray = True
            Exit For
        End If
    Next

End Function

Function IsInCollection(arr As Collection, valueToCheck As String) As Boolean
    IsInCollection = False
    For Each n In arr
        If n = valueToCheck Then
            IsInCollection = True
            Exit For
        End If
    Next

End Function

Function IsInShapes(arr As Shapes, valueToCheck As String) As Boolean
    IsInShapes = False
    For Each s In arr
        If s.Name = valueToCheck Then
            IsInShapes = True
            Exit For
        End If
    Next

End Function

Function IsInShapeRange(arr As ShapeRange, valueToCheck As String) As Boolean
    IsInShapeRange = False
    For Each s In arr
        If s.Name = valueToCheck Then
            IsInShapeRange = True
            Exit For
        End If
    Next

End Function

Sub ClearCollection(arr As Collection)
    Set arr = New Collection
End Sub

Sub AddAll(c_out As Collection, c2 As Collection)
    For Each Item In c2
      c_cout.Add Item
    Next Item
End Sub

Function ToCollection(a As Variant) As Collection
    Dim c As New Collection
    For Each Item In a
      c.Add Item
    Next
    Set ToCollection = c
End Function

Function ToArray(c As Collection) As Variant
    Dim a() As Variant
    ReDim a(1 To c.Count)
    For i = 1 To c.Count
      a(i) = c(i)
    Next
    ToArray = a
End Function

Function RunCmd(strCMD As String, Optional waitOnReturn As Boolean = True, Optional windowStyle As Integer = 1)
    '若無法執行，須引用 "Windows Script Host Object Model"
    ' (工具 > 設定引用項目 >勾選)
    ' 使用 WScript.Shell 方式
    ' 參數：
    ' strCMD 執行字串
    ' windowStyle 視窗樣式，1為顯示 0不顯示
    ' waitOnReturn 是否等待返回
    
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim errorCode As Integer
    On Error GoTo ErrZone
    errorCode = wsh.Run(strCMD, windowStyle, waitOnReturn)

    If errorCode = 0 Then
 '       MsgBox "OK!"
    Else
        MsgBox "執行錯誤" & vbCrLf & "代碼：" & errorCode & vbCrLf & "執行程式：" & strCMD
        Exit Function
    End If
    Exit Function
    
ErrZone:
    MsgBox "WScript.Shell發生錯誤：" & vbCrLf & Err.Number & ":" & Err.Description
Resume Next

End Function
