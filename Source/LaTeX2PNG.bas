Attribute VB_Name = "LaTeX2PNG"
Sub Code2PNG(oldShape As Shape, TempDir As String, FilePrefix As String, code As String)
    ' Run LaTeX Editor
    'TODO: [FIXME] 換Word、PowerPoint、Excel檔名時，執行會發生錯誤
    
    Dim sourceFileName As String, pictureFileName As String
    
    TempDir = "C:\Temp\" 'TODO: need to be portable to Mac OS X
    
    
    '[DONE] Use generateLaTaXName to generate FilePrefix
    If selectionIsLaTeXShape() Then
        FilePrefix = oldShape.Name
    Else
        FilePrefix = generateLaTeXName()
    End If
    
    Debug.Assert StartWith(FilePrefix, "tex4office_obj")
    
    
    
    sourceFileName = FilePrefix & ".tex"
    pictureFileName = FilePrefix & ".png"
    
    ' If TempDir doesn't exist, create one
    If Dir(TempDir, vbDirectory) = "" Then
       MkDir TempDir
    End If
    
    ' Setup work directory
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ChDir Environ("APPDATA") & "\Microsoft\AddIns\TeX4Office_Editor\"
    
    'Delete tex file
    If Dir(TempDir & sourceFileName) <> Empty Then
        fs.DeleteFile TempDir & sourceFileName
    End If
    
    'Delete png file
    If Dir(TempDir & pictureFileName) <> Empty Then
        fs.DeleteFile TempDir & pictureFileName
    End If
    

    If selectionIsLaTeXShape() Then
        code = oldShape.AlternativeText
        WriteLaTeXToFile code, TempDir, FilePrefix
    End If
    
    RunCmd "TeX4Office_WindowsFormsApplication.exe  " & TempDir & sourceFileName, True, vbNormalFocus
    
    
    If Dir(TempDir & pictureFileName) = Empty Then
        Exit Sub
    End If
    
    ReadLaTeXFromFile code, TempDir, FilePrefix
End Sub
