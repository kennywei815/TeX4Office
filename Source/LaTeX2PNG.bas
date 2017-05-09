Attribute VB_Name = "LaTeX2PNG"
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
'                            LaTeX2PNG Module
'             Interact with the LaTeX editor for the Main module
'*****************************************************************************




Sub LaTeX2PNG_Func(oldShape As Shape, TempDir As String, FilePrefix As String, code As String, Optional BatchMode As Boolean = False)
    ' Run LaTeX Editor
    
    '[TODO]: [FIXME] Error occurs when we change the file name of the Word, Excel, or PowerPoint file.
    
    Dim sourceFileName As String, pictureFileName As String, sourceFilePath As String, pictureFilePath As String
    
    TempDir = "C:\Temp\" '[TODO]: need to be portable to Mac OS X
    
    
    '==============================================================================================================================================
    ' Step 1:  Use generateLaTaXName to generate FilePrefix
    '==============================================================================================================================================
    If selectionIsLaTeXShape() Then
        FilePrefix = oldShape.Name
    Else
        FilePrefix = generateLaTeXName()
    End If
    
    Debug.Assert StartWith(FilePrefix, "tex4office_obj")
    
    
    
    sourceFileName = FilePrefix & ".tex"
    pictureFileName = FilePrefix & ".png"
    
    sourceFilePath = TempDir & FilePrefix & ".tex"
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
    
    ChDir Environ("APPDATA") & "\Microsoft\AddIns\TeX4Office_Editor\"
    
    '==============================================================================================================================================
    ' Step 4: Delete old files
    '==============================================================================================================================================
    ' Delete tex file
    If Dir(sourceFilePath) <> Empty Then
        fs.DeleteFile sourceFilePath
    End If
    
    'Delete png file
    If Dir(pictureFilePath) <> Empty Then
        fs.DeleteFile pictureFilePath
    End If
    

    '==============================================================================================================================================
    ' Step 5: Write LaTeX code from oldShape to tex file, and open it in TeX4Office editor
    '==============================================================================================================================================
    If selectionIsLaTeXShape() Then
        code = oldShape.AlternativeText
        WriteToFile_UTF8 code, sourceFilePath
    End If
    
    If BatchMode Then
        RunCmd "TeX4Office_Editor.exe  --batch-mode " & sourceFilePath, True, vbNormalFocus
        'RunCmd "TeX4Office_Editor.exe  " & sourceFilePath, True, vbNormalFocus
        'RunCmd "TeX4Office_WindowsFormsApplication.exe  --batch-mode " & sourceFilePath, True, vbNormalFocus
    Else
        RunCmd "TeX4Office_Editor.exe  " & sourceFilePath, True, vbNormalFocus
        'RunCmd "TeX4Office_WindowsFormsApplication.exe  " & sourceFilePath, True, vbNormalFocus
    End If
    
    
    '==============================================================================================================================================
    ' Step 6: If the user chooses not to generate new PNG file, exit the subroutine
    '==============================================================================================================================================
    If Dir(pictureFilePath) = Empty Then
        Exit Sub
    End If
    
    '==============================================================================================================================================
    ' Step 7: Read LaTeX code from tex file
    '==============================================================================================================================================
    ReadFromFile_UTF8 code, sourceFilePath
End Sub
