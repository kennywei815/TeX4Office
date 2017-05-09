Attribute VB_Name = "Reload_Build_Module"
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
'                         Reload_Build_Module Module
'               help developers reload Build module from Build.bas
'*****************************************************************************




'==============================================================================================================================================
'                                            Platform Constants
'==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = Word
Private Const PLATFORM_DEF As String = "#Const PLATFORM = Word"


Public WorkDir As String
Public VBComponents
Public Sep As String


'==============================================================================================================================================
'                                            Common_Vars Function
'                                        Set up commonly used variables
'==============================================================================================================================================
Sub Common_Vars()

#If PLATFORM = PowerPoint Then
    WorkDir = ActivePresentation.path
    Set VBComponents = ActivePresentation.VBProject.VBComponents
            
#ElseIf PLATFORM = Word Then
    WorkDir = ActiveDocument.path
    Set VBComponents = ActiveDocument.VBProject.VBComponents
            
#ElseIf PLATFORM = Excel Then
    WorkDir = ActiveWorkbook.path
    Set VBComponents = ActiveWorkbook.VBProject.VBComponents
    
#End If
    
    'Sep = Application.PathSeparator
    Sep = "\"
    
    ChDir WorkDir

End Sub


'==============================================================================================================================================
'                                            Reload Modules from .bas
'==============================================================================================================================================
Sub Reload_Build_Module()

    Common_Vars
    
    Set fs = CreateObject("Scripting.FileSystemObject")

    With VBComponents
    
        .Remove .Item("Build")
        .Import WorkDir & Sep & "Build.bas"
        
        
        .Item("Build").CodeModule.ReplaceLine 48, PLATFORM_DEF
        .Item("Build").CodeModule.ReplaceLine 49, "Private Const PLATFORM_DEF As String = """ & PLATFORM_DEF & """"
        
    End With

End Sub
