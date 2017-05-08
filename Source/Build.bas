Attribute VB_Name = "Build"
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = PowerPoint
Const PLATFORM_DEF As String = "#Const PLATFORM = PowerPoint"


#Const TeX4Office = 0
#Const ImportImage = 1

#Const PROGRAM = ImportImage
Const PROGRAM_DEF As String = "#Const PROGRAM = ImportImage"


'''''==============================================================================================================================================
'''''                                            Build Functions
'''''==============================================================================================================================================
Sub Build()

#If PLATFORM = PowerPoint Then
    WorkDir = ActivePresentation.Path
            
#ElseIf PLATFORM = Word Then
    WorkDir = ActiveDocument.Path
            
#ElseIf PLATFORM = Excel Then
    WorkDir = ActiveWorkbook.Path
    
#End If
    
    'Sep = Application.PathSeparator
    Sep = "\"


#If PROGRAM = TeX4Office Then
    fn = "TeX4Office_v0_2Beta"
    fn_addin = "TeX4Office_Beta1"

#ElseIf PROGRAM = ImportImage Then
    fn = "ImportImagePlus_v0_2Beta"
    fn_addin = "ImportImagePlus_Beta1"
        
#End If
    
    f_path = WorkDir & Sep & fn
    f_path_addin = WorkDir & Sep & fn_addin
    addinDir = WorkDir & Sep & "Installer\AddIns" & Sep
    wordStartDir = WorkDir & Sep & "Installer\Word\STARTUP" & Sep
    excelStartDir = WorkDir & Sep & "Installer\Excel\XLSTART" & Sep
    
    'MsgBox f_path
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    ChDir WorkDir

#If PLATFORM = PowerPoint Then
    With ActivePresentation
        .Save
        .SaveAs f_path, ppSaveAsOpenXMLAddin
        '.SaveAs f_path, ppSaveAsAddin
        .SaveAs f_path, ppSaveAsOpenXMLPresentationMacroEnabled
    End With
    
    If fs.FileExists(f_path_addin & ".ppam") Then
        fs.DeleteFile f_path_addin & ".ppam"
    End If
    If fs.FileExists(addinDir & fn_addin & ".ppam") Then
        fs.DeleteFile addinDir & fn_addin & ".ppam"
    End If
    
    fs.MoveFile f_path & ".ppam", f_path_addin & ".ppam"
    fs.MoveFile f_path_addin & ".ppam", addinDir
            
#ElseIf PLATFORM = Word Then
    With ActiveDocument
        .Save
        .SaveAs f_path, wdFormatXMLTemplateMacroEnabled
        '.SaveAs f_path, wdFormatTemplate
        .SaveAs f_path, wdFormatXMLDocumentMacroEnabled
    End With
    
    If fs.FileExists(f_path_addin & ".dotm") Then
        fs.DeleteFile f_path_addin & ".dotm"
    End If
    If fs.FileExists(addinDir & fn_addin & ".dotm") Then
        fs.DeleteFile addinDir & fn_addin & ".dotm"
    End If
    If fs.FileExists(wordStartDir & fn_addin & ".dotm") Then
        fs.DeleteFile wordStartDir & fn_addin & ".dotm"
    End If
    
    fs.MoveFile f_path & ".dotm", f_path_addin & ".dotm"
    fs.CopyFile f_path_addin & ".dotm", wordStartDir
    fs.MoveFile f_path_addin & ".dotm", addinDir
            
#ElseIf PLATFORM = Excel Then
    With ActiveWorkbook
        .Save
        .SaveAs f_path, xlOpenXMLAddIn
        '.SaveAs f_path, xlAddIn
        .SaveAs f_path, xlOpenXMLWorkbookMacroEnabled
    End With
    
    If fs.FileExists(f_path_addin & ".xlam") Then
        fs.DeleteFile f_path_addin & ".xlam"
    End If
    If fs.FileExists(addinDir & fn_addin & ".xlam") Then
        fs.DeleteFile addinDir & fn_addin & ".xlam"
    End If
    If fs.FileExists(excelStartDir & fn_addin & ".xlam") Then
        fs.DeleteFile excelStartDir & fn_addin & ".xlam"
    End If
    
    fs.MoveFile f_path & ".xlam", f_path_addin & ".xlam"
    fs.CopyFile f_path_addin & ".xlam", excelStartDir
    fs.MoveFile f_path_addin & ".xlam", addinDir
            
#End If


End Sub



'''''==============================================================================================================================================
'''''                                            Reload Modules from .bas
'''''==============================================================================================================================================
Sub Reload_Modules()

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
        .Remove .Item("AutoRun")
        .Remove .Item("Common_Utilities")
        
#If PROGRAM = TeX4Office Then
        .Remove .Item("LaTeX2PNG")
        
#ElseIf PROGRAM = ImportImage Then
        .Remove .Item("Image2PNG")

#End If
        
        .Remove .Item("Main")
        .Remove .Item("Main_Helpers")
        
        
        .Import WorkDir & Sep & "AutoRun.bas"
        .Import WorkDir & Sep & "Common_Utilities.bas"

#If PROGRAM = TeX4Office Then
        .Import WorkDir & Sep & "LaTeX2PNG.bas"

#ElseIf PROGRAM = ImportImage Then
        .Import WorkDir & Sep & "Image2PNG.bas"
        
#End If
        
        .Import WorkDir & Sep & "Main.bas"
        .Import WorkDir & Sep & "Main_Helpers.bas"
        
        
        .Item("AutoRun").CodeModule.ReplaceLine 8, PLATFORM_DEF
        .Item("AutoRun").CodeModule.ReplaceLine 14, PROGRAM_DEF
        
        .Item("Main").CodeModule.ReplaceLine 8, PLATFORM_DEF
        .Item("Main").CodeModule.ReplaceLine 14, PROGRAM_DEF
        
        .Item("Main_Helpers").CodeModule.ReplaceLine 8, PLATFORM_DEF
        .Item("Main_Helpers").CodeModule.ReplaceLine 14, PROGRAM_DEF
        
    End With

End Sub
