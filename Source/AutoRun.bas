Attribute VB_Name = "AutoRun"
'''''==============================================================================================================================================
'''''                                            Platform-dependent Macros
'''''==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = Word


#Const TeX4Office = 0
#Const ImportImage = 1

#Const PROGRAM = TeX4Office


'''''==============================================================================================================================================
'''''                                            AutoRun Macros (executed when the application starts)
'''''==============================================================================================================================================
' Runs when the add-in is loaded

Sub AutoExec()
    InitializeApp
End Sub

Sub AutoExit()
    UnInitializeApp
End Sub

Sub Auto_Open()
    InitializeApp
End Sub

Sub Auto_Close()
    UnInitializeApp
End Sub


Sub InitializeApp()
    'Set theAppEventHandler.App = Application
#If PROGRAM = TeX4Office Then
    AddMenuItem "New/Edit LaTeX Display...", "EditLaTeX", 18 '226
    
#ElseIf PROGRAM = ImportImage Then
    AddMenuItem "Insert/Change Image...", "InsertPicture", 18 '226
    
#End If
    
End Sub

Sub UnInitializeApp()
    
#If PROGRAM = TeX4Office Then
    RemoveMenuItem "New/Edit LaTeX Display..."
    
#ElseIf PROGRAM = ImportImage Then
    RemoveMenuItem "Insert/Change Image..."
    
#End If
    
End Sub






