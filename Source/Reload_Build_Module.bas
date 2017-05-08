Attribute VB_Name = "Reload_Build_Module"
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = Word
Const PLATFORM_DEF As String = "#Const PLATFORM = Word"


#Const TeX4Office = 0
#Const ImportImage = 1

#Const PROGRAM = TeX4Office
Const PROGRAM_DEF As String = "#Const PROGRAM = TeX4Office"


'''''==============================================================================================================================================
'''''                                            Reload Modules from .bas
'''''==============================================================================================================================================
Sub Reload_Build_Module()

#If PLATFORM = PowerPoint Then
    WorkDir = ActivePresentation.Path
            
#ElseIf PLATFORM = Word Then
    WorkDir = ActiveDocument.Path
            
#ElseIf PLATFORM = Excel Then
    WorkDir = ActiveWorkbook.Path
    
#End If
    
    'Sep = Application.PathSeparator
    Sep = "\"
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ChDir WorkDir


#If PLATFORM = PowerPoint Then
    With ActivePresentation.VBProject.VBComponents
            
#ElseIf PLATFORM = Word Then
    With ActiveDocument.VBProject.VBComponents
            
#ElseIf PLATFORM = Excel Then
    With ActiveWorkbook.VBProject.VBComponents
            
#End If
        '.Remove .Item("Build")
        .Import WorkDir & Sep & "Build.bas"
        
        
        .Item("Build").CodeModule.ReplaceLine 5, PLATFORM_DEF
        .Item("Build").CodeModule.ReplaceLine 6, "Const PLATFORM_DEF As String = """ & PLATFORM_DEF & """"
        .Item("Build").CodeModule.ReplaceLine 12, PROGRAM_DEF
        .Item("Build").CodeModule.ReplaceLine 13, "Const PPROGRAM_DEF As String = """ & PPROGRAM_DEF & """"
        
    End With

End Sub
