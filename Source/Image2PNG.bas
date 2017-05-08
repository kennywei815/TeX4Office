Attribute VB_Name = "Image2PNG"
Sub Code2PNG(oldShape As Shape, TempDir As String, FilePrefix As String, code As String)
    '[TODO]: refactor
    '[TODO]: generate code when done
    
    Dim sourceFileName As String, pictureFileName As String, Command As String
    
    
    TempDir = "C:\Temp\" '[TODO]: need to be portable to Mac OS X
    
    '[DONE] Use generateLaTaXName to generate FilePrefix
    If selectionIsLaTeXShape() Then
        FilePrefix = oldShape.Name
    Else
        FilePrefix = generateLaTeXName()
    End If
    
    Debug.Assert StartWith(FilePrefix, "importImage_plus_obj")
    
    pictureFileName = FilePrefix & ".png"
    
    ' If TempDir doesn't exist, create one
    If Dir(TempDir, vbDirectory) = "" Then
       MkDir TempDir
    End If
    
    ' Setup work directory
    Set fs = CreateObject("Scripting.FileSystemObject")

    'Delete png file
    If Dir("tex_file.png") <> Empty Then
        fs.DeleteFile "tex_file.png"
    End If
    
    ' Open the file dialog
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        If .Show = -1 Then
 
            ' Display paths of each file selected
            For lngCount = 1 To .SelectedItems.Count
            
                Command = "cmd /C convert   -units PixelsPerInch  -density 1200 -resize 1200x1200 " & QuotedStr(.SelectedItems(lngCount)) & " " & QuotedStr(TempDir & pictureFileName)
                
                'TODO: fix RunCmd
                'TODO: 顯示哪些格式不支援
                'Shell Command, vbMaximizedFocus
                RunCmd Command, True, vbMaximizedFocus
            Next lngCount
            
        End If
 
    End With
    
    
    'TODO: fix RunCmd & Turn on these code
    If Dir(TempDir & pictureFileName) = Empty Then
        Exit Sub
    End If
    
End Sub

