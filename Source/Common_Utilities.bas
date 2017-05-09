Attribute VB_Name = "Common_Utilities"
' 2017(c) TeX4Office
' Developer: Cheng-Kuan Wei
' URL: https://github.com/kennywei815/tex4office
'
' Licensed to the Apache Software Foundation (ASF) under one
' or more contributor license agreements.  See the NOTICE file
' distributed with this work for additional information
' regarding copyright ownership.  The ASF licenses this file
' to you under the Apache License, Version 2.0 (the
' "License"); you may not use this file except in compliance
' with the License.  You may obtain a copy of the License at
'
'   http://www.apache.org/licenses/LICENSE-2.0
'
' This file incorporates work from Jonathan Le Roux and Zvika
' Ben-Haim's IguanaTeX project which is originally released
' under the Creative Commons Attribution 3.0 License; you may
' not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'   https://creativecommons.org/licenses/by/3.0/
'
' Unless required by applicable law or agreed to in writing,
' software distributed under the License is distributed on an
' "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY
' KIND, either express or implied.  See the License for the
' specific language governing permissions and limitations
' under the License.




'*****************************************************************************
'                         Common_Utilities Module
'                  Implements some commonly used utilities
'*****************************************************************************




'==============================================================================================================================================
'                                            Platform-independent Functions
'==============================================================================================================================================
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

Sub ReadFromFile_UTF8(code As String, path As String)
    Const ForReading = 1, ForWriting = 2, ForAppending = 3
    
    'UseUTF8
    Dim objStream, strData

    Set objStream = CreateObject("ADODB.Stream")
    
    objStream.Charset = "utf-8"
    objStream.Open
    objStream.LoadFromFile (path)
    
    code = objStream.ReadText()
    
    objStream.Close
    Set objStream = Nothing
End Sub

Sub WriteToFile_UTF8(code As String, path As String)
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
        
        ' Workaround to avoid BOM in file:
        .Position = 3 'skip BOM
        .CopyTo BinaryStream
        .Flush
        .Close
    End With
    BinaryStream.SaveToFile path, 2 'Save binary data To disk
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


Function PackToHeader(code As String) As String

    Lines = Split(code, vbNewLine)
    
    PackToHeader = "%%% HEADER %%%" & vbNewLine
    
    For i = 1 To Lines.Count
        PackToHeader = PackToHeader & "%" & Lines(i) & vbNewLine
    Next
    
    PackToHeader = PackToHeader & "%%% END HEADER %%%"
    
End Function


Function UnpackFromHeader(code As String) As String

    Paragraph = Split(code, "%%% END HEADER %%%")
    header = Replace(Paragraph(1), "%%% HEADER %%%", "")
    
    Lines = Split(header, vbNewLine)
    UnpackFromHeader = ""
    
    For i = 1 To Lines.Count
        UnpackFromHeader = UnpackFromHeader & Replace(Expression:=Lines(i), Find:="%", Replace:="", Count:=1) & vbNewLine
    Next
    
End Function


Function ReadConfig(code As String, engine As String, dpi As Integer, fileName As String)
    Lines = Split(code, vbNewLine)
    
    For i = 1 To Lines.Count
        Key_Value = Split(Lines(i), ",")
        
        Key = Key_Value(1)
        Value = Key_Value(2)
        
        Select Case Key
           Case "engine"
              engine = Value
           Case "dpi"
              dpi = CInt(Value)
           Case "fileName"
              fileName = Value
           'Case Else
           
        End Select
    Next
End Function


Function WriteConfig(code As String, engine As String, dpi As Integer, fileName As String)
    
    code = "engine," & engine & vbNewLine _
         & "dpi," & dpi & vbNewLine _
         & "fileName," & fileName & vbNewLine
    
End Function


Function RunCmd(strCMD As String, Optional waitOnReturn As Boolean = True, Optional windowStyle As Integer = 1)
    ' Prerequisite: need to import "Windows Script Host Object Model"  (in "Tools" -> "References...")
    '
    ' using WScript.Shell
    '
    ' Arguments:
    '   strCMD:       command to be executed
    '   windowStyle:  window style: 1 means showing-up, while 0 means hiding
    '   waitOnReturn: wait for the command
    
    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")
    Dim errorCode As Integer
    On Error GoTo ErrZone
    errorCode = wsh.Run(strCMD, windowStyle, waitOnReturn)

    If errorCode = 0 Then
 '       MsgBox "OK!"
    Else
        MsgBox "Error" & vbCrLf & "Code¡G" & errorCode & vbCrLf & "Command¡G" & strCMD
        Exit Function
    End If
    Exit Function
    
ErrZone:
    MsgBox "WScript.Shell encounters error¡G" & vbCrLf & Err.Number & ":" & Err.Description
Resume Next

End Function
