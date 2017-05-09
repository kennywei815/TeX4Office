Attribute VB_Name = "AutoRun"
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
'                              AutoRun Module
'       Functions executed when Office starts. Mainly related to the UI.
'*****************************************************************************




'==============================================================================================================================================
'                                            Platform Constants
'==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = Word


'==============================================================================================================================================
'                                            AutoRun Macros (executed when the application starts)
'==============================================================================================================================================
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
    'AddMenuItem "New/Edit LaTeX Display...", "EditLaTeX", 18 '226
    'AddMenuItem "Insert/Change Image...", "InsertPicture", 37 '226
    
    
    AddMenuItem "New/Edit LaTeX Display...", "EditLaTeX", 1 '226
    AddMenuItem "Insert/Change Image...", "InsertPicture", 1 '226
    
End Sub

Sub UnInitializeApp()
    
    RemoveMenuItem "New/Edit LaTeX Display..."
    RemoveMenuItem "Insert/Change Image..."
    
End Sub






