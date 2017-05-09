Attribute VB_Name = "Image2PNG"
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
'                            Image2PNG Module
'             Load image and convert it to PNG for the Main module
'*****************************************************************************




Sub Image2PNG_Func(oldShape As Shape, TempDir As String, FilePrefix As String, code As String)
    '[TODO]: save code when done
    
    Dim sourceFileName As String, pictureFileName As String, pictureFilePath As String, Command As String
    
    
    TempDir = "C:\Temp\" '[TODO]: need to be portable to Mac OS X
    
    '==============================================================================================================================================
    ' Step 1:  Use generateLaTaXName to generate FilePrefix
    '==============================================================================================================================================
    If selectionIsImagePlusShape() Then
        FilePrefix = oldShape.Name
    Else
        FilePrefix = generateImagePlusName()
    End If
    
    Debug.Assert StartWith(FilePrefix, "importImage_plus_obj")
    
    pictureFileName = FilePrefix & ".png"
    pictureFilePath = TempDir & FilePrefix & ".png"
    
    '==============================================================================================================================================
    ' Step 2:  If TempDir doesn't exist, create one
    '==============================================================================================================================================
    If Dir(TempDir, vbDirectory) = "" Then
       MkDir TempDir
    End If
    
    '==============================================================================================================================================
    ' Step 3:  Setup work directory
    '==============================================================================================================================================
    Set fs = CreateObject("Scripting.FileSystemObject")

    '==============================================================================================================================================
    ' Step 4: Delete old files
    '==============================================================================================================================================
    'Delete png file
    If Dir("tex_file.png") <> Empty Then
        fs.DeleteFile "tex_file.png"
    End If
    
    '==============================================================================================================================================
    ' Step 5: Open the file dialog, and convert loaded image file into PNGs using convert.exe
    '==============================================================================================================================================
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        If .Show = -1 Then
 
            ' Display paths of each file selected
            For lngCount = 1 To .SelectedItems.Count
            
                Command = "cmd /C " & Environ("APPDATA") & "\Microsoft\AddIns\TeX4Office_Editor\ImageMagick-portable\" & "convert   -units PixelsPerInch  -density 1200 -resize 1200x1200 " & QuotedStr(.SelectedItems(lngCount)) & " " & QuotedStr(pictureFilePath)
                
                '[TODO]: Show all supported input file format of convert.exe
                RunCmd Command, True, vbNormalFocus
            Next lngCount
            
        End If
 
    End With
    
    
    '==============================================================================================================================================
    ' Step 6: If the user chooses to cancel the file dialog, exit the subroutine
    '==============================================================================================================================================
    If Dir(pictureFilePath) = Empty Then
        Exit Sub
    End If
    
End Sub

