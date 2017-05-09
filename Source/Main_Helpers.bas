Attribute VB_Name = "Main_Helpers"
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
'                           Main_Helpers Module
'                    Helper functions for the Main module
'*****************************************************************************




'==============================================================================================================================================
'                                            Platform Constants
'==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = Word


Const TeX4Office As Integer = 0
Const ImportImage As Integer = 1



Public Const ToBeRegroupedLevel As Integer = 1
Public Const InvalidLevel As Integer = -100
Public Const InvalidName As String = "!@#$#%^$&^%*&&^*(*&)"


#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As LongPtr) 'MS Office 64 Bit
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal milliseconds as Long)            'MS Office 32 Bit
#End If


' Get current slide or document
Public AllShapes As Shapes

#If PLATFORM = PowerPoint Then
    Public SlideIndex As Long
    Public sld As Slide
    Public osld As Slide
            
#ElseIf PLATFORM = Word Then
    Public sld As Document
    Public osld As Document
            
#ElseIf PLATFORM = Excel Then
    Public sld As Worksheet
    Public osld As Worksheet
            
#End If


'==============================================================================================================================================
'                                            Plateform-dependent Helper Functions
'==============================================================================================================================================
Function selectionIsGroup() As Boolean
    Set Sel = ActiveWindow.Selection
    
#If PLATFORM = PowerPoint Then
    selectionIsGroup = (Sel.ShapeRange.Type = msoGroup And Sel.hasChildShapeRange)

#ElseIf PLATFORM = Word Then
    '[Known Issue]: In Word, Sel.HasChildShapeRange & Sel.ShapeRange.GroupItems does not work properly.
    selectionIsGroup = (Sel.ShapeRange.Type = msoGroup And Sel.ChildShapeRange.Count >= 1)

#ElseIf PLATFORM = Excel Then
    '[Known Issue]: In Excel, we don't have Selection type, and we don't know how to select current selected shape's group.
    selectionIsGroup = False
            
#End If

End Function


Function selectionIsLaTeXShape() As Boolean
    Set Sel = ActiveWindow.Selection
    
    Dim isShape As Boolean



#If PLATFORM = PowerPoint Then
    isShape = (Sel.Type = ppSelectionShapes)
    
    If isShape Then
        If Sel.ShapeRange.Type = msoPicture Then
            selectionIsLaTeXShape = True
            
        ElseIf Sel.ShapeRange.Type = msoGroup And Sel.hasChildShapeRange Then
            selectionIsLaTeXShape = isLaTeXShape(Sel.ChildShapeRange(1)) 'When selection is a Group or an item in a Group, we need to use isLaTeXShape() to determine whether selection is a LaTeX shape.
            
        Else
            selectionIsLaTeXShape = False
        End If
    Else
        selectionIsLaTeXShape = False
    End If



#ElseIf PLATFORM = Word Then
    isShape = (Sel.Type = wdSelectionShape)
    
    If isShape Then
        If Sel.ShapeRange.Type = msoPicture Then
            selectionIsLaTeXShape = True

        ElseIf Sel.ShapeRange.Type = msoGroup And Sel.ChildShapeRange.Count >= 1 Then  '[Known Issue]: In Word, Sel.HasChildShapeRange & Sel.ShapeRange.GroupItems does not work properly.
            selectionIsLaTeXShape = isLaTeXShape(Sel.ChildShapeRange(1)) 'When selection is a Group or an item in a Group, we need to use isLaTeXShape() to determine whether selection is a LaTeX shape.
            
        Else
            selectionIsLaTeXShape = False
        End If
    Else
        selectionIsLaTeXShape = False
    End If



#ElseIf PLATFORM = Excel Then
    '[NOTE] In Excel, we don't have Selection type, so we need to use TypeName(Sel) & TypeOf Sel to get the type information of current selection.
    
    '====================================================================================== BEGIN DEBUG ======================================================================================
    'MsgBox "TypeName(Sel)=" & TypeName(Sel)
    'Single shape:  Picture
    'In a group:    Picture
    'Group:         GroupObject
    '====================================================================================== END DEBUG ======================================================================================
    
    isShape = (TypeOf Sel Is Picture)
    
    '====================================================================================== BEGIN DEBUG ======================================================================================
    'MsgBox "isShape=" & isShape
    'MsgBox "Sel.ShapeRange.Count=" & Sel.ShapeRange.Count
    ''MsgBox "Sel.hasChildShapeRange=" & Sel.hasChildShapeRange
    ''MsgBox "Sel.ChildShapeRange.Count=" & Sel.ChildShapeRange.Count
    
    'MsgBox "TypeName(Sel.ShapeRange(1))=" & TypeName(Sel.ShapeRange(1))
    'Single shape:  Shape
    'In a group:    Shape
    'Group:         Shape
        
    'MsgBox "Sel.ShapeRange(1).Type=" & Sel.ShapeRange(1).Type
    '***** Same as in PowerPoint & Word *****
    'Single shape:  msoPicture
    'In a group:    msoPicture
    'Group:         msoGroup
    '====================================================================================== END DEBUG ======================================================================================
    
    If isShape Then
        selectionIsLaTeXShape = isLaTeXShape(Sel.ShapeRange(1))
    Else
        selectionIsLaTeXShape = False
    End If


#End If

End Function


Function isLaTeXShape(s As Shape) As Boolean
    '[DONE] check if oldShape.Name has .tex suffix or oldShape.AlternativeText contains tex code
    isLaTeXShape = (StartWith(s.Name, "tex4office_obj") And s.AlternativeText <> "")
    
End Function


Function generateLaTeXName() As String
    Const RAND_MAX As Integer = 32767

    Dim id As Integer
    
    Do
        id = Int((RAND_MAX * Rnd()) + 1)                   ' Generate random value between 1 and RAND_MAX+1.
        generateLaTeXName = "tex4office_obj" & id

    Loop While IsInShapes(AllShapes, generateLaTeXName)
    
End Function


Function selectionIsImagePlusShape() As Boolean
    Set Sel = ActiveWindow.Selection
    
    Dim isShape As Boolean



#If PLATFORM = PowerPoint Then
    isShape = (Sel.Type = ppSelectionShapes)
    
    If isShape Then
        If Sel.ShapeRange.Type = msoPicture Then
            selectionIsImagePlusShape = True
            
        ElseIf Sel.ShapeRange.Type = msoGroup And Sel.hasChildShapeRange Then
            selectionIsImagePlusShape = isImagePlusShape(Sel.ChildShapeRange(1)) 'When selection is a Group or an item in a Group, we need to use isImagePlusShape() to determine whether selection is a ImagePlus shape.
            
        Else
            selectionIsImagePlusShape = False
        End If
    Else
        selectionIsImagePlusShape = False
    End If



#ElseIf PLATFORM = Word Then
    isShape = (Sel.Type = wdSelectionShape)
    
    If isShape Then
        If Sel.ShapeRange.Type = msoPicture Then
            selectionIsImagePlusShape = True

        ElseIf Sel.ShapeRange.Type = msoGroup And Sel.ChildShapeRange.Count >= 1 Then  '[Known Issue]: In Word, Sel.HasChildShapeRange & Sel.ShapeRange.GroupItems does not work properly.
            selectionIsImagePlusShape = isImagePlusShape(Sel.ChildShapeRange(1)) 'When selection is a Group or an item in a Group, we need to use isImagePlusShape() to determine whether selection is a ImagePlus shape.
            
        Else
            selectionIsImagePlusShape = False
        End If
    Else
        selectionIsImagePlusShape = False
    End If



#ElseIf PLATFORM = Excel Then
    '[NOTE] In Excel, we don't have Selection type, so we need to use TypeName(Sel) & TypeOf Sel to get the type information of current selection.
    
    '====================================================================================== BEGIN DEBUG ======================================================================================
    'MsgBox "TypeName(Sel)=" & TypeName(Sel)
    'Single shape:  Picture
    'In a group:    Picture
    'Group:         GroupObject
    '====================================================================================== END DEBUG ======================================================================================
    
    isShape = (TypeOf Sel Is Picture)
    
    '====================================================================================== BEGIN DEBUG ======================================================================================
    'MsgBox "isShape=" & isShape
    'MsgBox "Sel.ShapeRange.Count=" & Sel.ShapeRange.Count
    ''MsgBox "Sel.hasChildShapeRange=" & Sel.hasChildShapeRange
    ''MsgBox "Sel.ChildShapeRange.Count=" & Sel.ChildShapeRange.Count
    
    'MsgBox "TypeName(Sel.ShapeRange(1))=" & TypeName(Sel.ShapeRange(1))
    'Single shape:  Shape
    'In a group:    Shape
    'Group:         Shape
        
    'MsgBox "Sel.ShapeRange(1).Type=" & Sel.ShapeRange(1).Type
    '***** Same as in PowerPoint & Word *****
    'Single shape:  msoPicture
    'In a group:    msoPicture
    'Group:         msoGroup
    '====================================================================================== END DEBUG ======================================================================================
    
    If isShape Then
        selectionIsImagePlusShape = isImagePlusShape(Sel.ShapeRange(1))
    Else
        selectionIsImagePlusShape = False
    End If


#End If

End Function


Function isImagePlusShape(s As Shape) As Boolean
    '[DONE] check if oldShape.Name has .tex suffix or oldShape.AlternativeText contains tex code
    isImagePlusShape = (StartWith(s.Name, "importImage_plus_obj") And s.AlternativeText <> "")
    
End Function


Function generateImagePlusName() As String
    Const RAND_MAX As Integer = 32767

    Dim id As Integer
    
    Do
        id = Int((RAND_MAX * Rnd()) + 1)                   ' Generate random value between 1 and RAND_MAX+1.
        generateImagePlusName = "importImage_plus_obj" & id

    Loop While IsInShapes(AllShapes, generateImagePlusName)
    
End Function


'==============================================================================================================================================
'                                            AddDisplayShape Function
'                        Add picture as shape taking care of not inserting it in empty placeholder
'==============================================================================================================================================
Function AddDisplayShape(PROGRAM As Integer, path As String, PosX As Single, PosY As Single) As Shape
' from http://www.vbaexpress.com/forum/showthread.php?47687-Addpicture-adds-the-picture-to-a-placeholder-rather-as-a-new-shape
' modified based on http://www.vbaexpress.com/forum/showthread.php?37561-Delete-empty-placeholders
    
    Dim oshp As Shape
    On Error Resume Next
    
    If Err <> 0 Then Exit Function
    On Error GoTo 0
    
#If PLATFORM = PowerPoint Or PLATFORM = Excel Then
    For Each oshp In osld.Shapes
        If oshp.Type = msoPlaceholder Then
            If oshp.PlaceholderFormat.ContainedType = msoAutoShape Then
                If oshp.HasTextFrame Then
                    If Not oshp.TextFrame.HasText Then oshp.TextFrame.TextRange = "DUMMY"
                End If
            End If
        End If
    Next oshp
    
#End If


If PROGRAM = TeX4Office Then
    Set AddDisplayShape = osld.Shapes.AddPicture(path, msoFalse, msoTrue, PosX, PosY, -1, -1)
    
ElseIf PROGRAM = ImportImage Then
    Set AddDisplayShape = osld.Shapes.AddPicture(path, msoFalse, msoTrue, PosX, PosY, 1200, 1200)
    
End If
    
#If PLATFORM = PowerPoint Or PLATFORM = Excel Then
    For Each oshp In osld.Shapes
        If oshp.Type = msoPlaceholder Then
            If oshp.PlaceholderFormat.ContainedType = msoAutoShape Then
                If oshp.HasTextFrame Then
                    If oshp.TextFrame.TextRange = "DUMMY" Then oshp.TextFrame.DeleteText
                End If
            End If
        End If
    Next oshp
    
#End If
    
End Function



'==============================================================================================================================================
'                                            MoveAnimation Function
'                                 Move the animation settings of oldShape to newShape
'==============================================================================================================================================
Sub MoveAnimation(oldShape As Shape, newShape As Shape)
    
#If PLATFORM = PowerPoint Then
    With osld.TimeLine
        Dim eff As Effect
        For Each eff In .MainSequence
            If eff.Shape.Name = oldShape.Name Then eff.Shape = newShape
        Next
    End With
    
#End If
    
End Sub


'==============================================================================================================================================
'                                            DeleteAnimation Function (Currently not used)
'                                              Delete the animation settings of oldShape
'==============================================================================================================================================
Sub DeleteAnimation(oldShape As Shape)
    
#If PLATFORM = PowerPoint Then
    With osld.TimeLine
        For i = .MainSequence.Count To 1 Step -1
            Dim eff As Effect
            Set eff = .MainSequence(i)
            If eff.Shape.Name = oldShape.Name Then eff.Delete
        Next
    End With
    
#End If
    
End Sub



'==============================================================================================================================================
'                                            Plateform-independent Helper Functions
'==============================================================================================================================================


Sub AddTagsToShape(vSh As Shape, code As String, FilePrefix As String)
    vSh.AlternativeText = code
    vSh.Name = FilePrefix
End Sub


'==============================================================================================================================================
'                                          RecordGroupHierarchy_and_Ungroup Function (for MS Word (without using CurShape.GroupItems()))
'                       Record group hierarchy information of current shape and recursively ungroup these groups
'==============================================================================================================================================
Function RecordGroupHierarchy_and_Ungroup(CurShape As Shape, TargetName As String, TargetSelectionName As String, ShapeNames As Collection, Layers As Collection, GroupNames As Collection) As Long
    ' This function expects to receive a grouped or ungrouped Shape (CurShape)
    ' We ungroup to reveal the structure at the layer below, and then regroup the groups which do not contain the target LaTeX shape.
    '
    ' Arguments:
    '   ShapeNames is the list of names of (leaf) elements in this group
    '   TargetName is the display which is being modified. We're going down the branch containing it.
    
    ''====================================================================================== BEGIN DEBUG ======================================================================================
    'MsgBox "In RecordGroupHierarchy_and_Ungroup()"
    'MsgBox "CurShape.Name=" & CurShape.Name
    ''====================================================================================== END DEBUG ======================================================================================
    
    Dim ToBeRegrouped_Now As Boolean
    ToBeRegrouped_Now = True
    
    
    Dim CurShapeName As String
    CurShapeName = CurShape.Name
    
    Dim IsGroup As Boolean
    IsGroup = (CurShape.Type = msoGroup)
    
    CurShape.Select
    Set Sel = ActiveWindow.Selection
    
    
    If Not IsGroup Then
        ' we reached the final layer
        
        RecordGroupHierarchy_and_Ungroup = 1
        
        ShapeNames.Add CurShapeName
        Layers.Add 1, CurShapeName
        GroupNames.Add CurShapeName, CurShapeName
        
    Else
        ' If received a grouped Shape, we ungroup it to reveal the structure at the layer below,
        ' since the element being edited is still within a group
        
        CurShape.Ungroup
        
        
        ' Build ShapeNames as names of the objects in the top-level group
        Dim RecordGroupHierarchy_and_Ungroup_In As Integer
        RecordGroupHierarchy_and_Ungroup_In = InvalidLevel
        
        Dim FoundTarget As Boolean
        FoundTarget = False
        
        For Each s In Sel.ShapeRange
            ''====================================================================================== BEGIN DEBUG ======================================================================================
            'MsgBox "s.Name=" & s.Name
            ''====================================================================================== END DEBUG ======================================================================================
        
            If s.Name <> TargetName Then
                ShapeNames.Add s.Name
                
            Else
                ToBeRegrouped_Now = False
                RecordGroupHierarchy_and_Ungroup_In = 1
                
                TargetSelectionName = CurShapeName  'means for future query when regrouping
                
                FoundTarget = True
                
            End If
        Next
        
        
        
        ' Reveal the structure at the layer below:
        ' Hierachically ungroup the groups at the layer below, and then regroup the groups which do not contain the target LaTeX shape.
        
        Dim ShapeNames_Tmp As New Collection  ' shapes in the temporary group
        Dim ShapeNames_In As New Collection   ' shapes in the same group
        Dim ShapeNames_Out As New Collection  ' shapes not in the same group
        Dim Layers_Tmp As New Collection
        Dim Layers_in As New Collection
        Dim Layers_Out As New Collection
        Dim GroupNames_Tmp As New Collection
        Dim GroupNames_in As New Collection
        Dim GroupNames_Out As New Collection


        If FoundTarget Then ' Optimization
            For Each n In ShapeNames
                ShapeNames_Out.Add n
            Next
        Else
        
            For Each n In ShapeNames
                Dim SelShape_Tmp As Shape
                Set SelShape_Tmp = AllShapes(n)
                
                Dim IsGroup_Tmp As Boolean
                IsGroup_Tmp = (SelShape_Tmp.Type = msoGroup)
                
                ClearCollection ShapeNames_Tmp
                ClearCollection Layers_Tmp
                ClearCollection GroupNames_Tmp
                
                MaxGroupLevel_Tmp = RecordGroupHierarchy_and_Ungroup(SelShape_Tmp, TargetName, TargetSelectionName, ShapeNames_Tmp, Layers_Tmp, GroupNames_Tmp)
                
                ' Check if SelShape_Tmp is a group and contains target shape (with return value MaxGroupLevel_Tmp=ToBeRegroupedLevel)
                'If MaxGroupLevel_Tmp <> ToBeRegroupedLevel Then  'Actually we need this if statement only, but these multiple levels of if statements help us debug easily
                If IsGroup_Tmp Then
                    If MaxGroupLevel_Tmp <> ToBeRegroupedLevel Then
                        ' Contains Target Shape (with TargetName)  =>  Add to ShapeNames_In¡BLayers_In¡BGroupNames_In
                        ToBeRegrouped_Now = False
                        RecordGroupHierarchy_and_Ungroup_In = MaxGroupLevel_Tmp
                        
                        For Each n_Tmp In ShapeNames_Tmp
                            ShapeNames_In.Add n_Tmp
                            Layers_in.Add Layers_Tmp(n_Tmp), n_Tmp
                            GroupNames_in.Add GroupNames_Tmp(n_Tmp), n_Tmp
                        Next
                    Else
                        ' Maintains current state of ToBeRegrouped_Now
                        ' Add to ShapeNames_Out¡BLayers_Out¡BGroupNames_Out
                        ShapeNames_Out.Add n
                        
                        Debug.Assert (MaxGroupLevel_Tmp = 1)
                        Debug.Assert (ShapeNames_Tmp.Count = 1 And ShapeNames_Tmp(1) = n)
                        Debug.Assert (Layers_Tmp.Count = 1)
                        Debug.Assert (GroupNames_Tmp.Count = 1)
                        
                    End If
                    
                Else
                    ' Add to ShapeNames_Out¡BLayers_Out¡BGroupNames_Out
                    ShapeNames_Out.Add n
                        
                    Debug.Assert (MaxGroupLevel_Tmp = 1)
                    Debug.Assert (ShapeNames_Tmp.Count = 1 And ShapeNames_Tmp(1) = n)
                    Debug.Assert (Layers_Tmp.Count = 1)
                    Debug.Assert (GroupNames_Tmp.Count = 1)
                    
                End If
            Next
        End If
            
            
        ' Record ShapeNames, Layers, and GroupNames
        
        ClearCollection ShapeNames
        
        
        ' For all elements in that group, tag them
        For Each n In ShapeNames_In
            ShapeNames.Add n
            Layers.Add Layers_in.Item(n), n
            GroupNames.Add GroupNames_in.Item(n), n
        Next
        
        ' For all elements not in that group, tag them
        If ToBeRegrouped_Now Then
            Debug.Assert (ShapeNames_In.Count = 0)
            Debug.Assert (Layers_in.Count = 0)
            Debug.Assert (GroupNames_in.Count = 0)
        
            Dim newGroup As Shape
            Set newGroup = sld.Shapes.Range(ToArray(ShapeNames_Out)).Group()
            newGroup.Name = CurShapeName
            
            ' We don't need the Child Names of ShapeNames_Out, Layers_Out, and GroupNames_Out.
            ShapeNames.Add newGroup.Name
            Layers.Add 1, newGroup.Name
            GroupNames.Add newGroup.Name, newGroup.Name
            
            Debug.Assert (RecordGroupHierarchy_and_Ungroup_In = InvalidLevel)
            
            RecordGroupHierarchy_and_Ungroup = 1
            
        Else
            ' Combine with ShapeNames_Out, Layers_Out, and GroupNames_Out
            For Each n In ShapeNames_Out
                ShapeNames.Add n
                Layers.Add RecordGroupHierarchy_and_Ungroup_In, n
                GroupNames.Add CurShapeName, n
            Next
            
            
            Debug.Assert (RecordGroupHierarchy_and_Ungroup_In >= 1)
            
            RecordGroupHierarchy_and_Ungroup = RecordGroupHierarchy_and_Ungroup_In + 1
            
        End If
    End If
    
    Debug.Assert (RecordGroupHierarchy_and_Ungroup >= 1)
    
End Function


'==============================================================================================================================================
'                                       RecordGroupHierarchy_and_Ungroup_Fast Function (for MS PowerPoint))
'                       Record group hierarchy information of current shape and recursively ungroup these groups
'==============================================================================================================================================
Function RecordGroupHierarchy_and_Ungroup_Fast(ShapeNames As Variant, TargetName As String, Layers As Collection, SelectionNames As Collection) As Long
    ' This function expects to receive a grouped ShapeRange (Sel.ShapeRange)
    ' We ungroup to reveal the structure at the layer below, and then regroup the groups which do not contain the target LaTeX shape.
    '
    ' Arguments:
    '   ShapeNames is the list of names of (leaf) elements in this group
    '   TargetName is the display which is being modified. We're going down the branch containing it.
    
    
    '[TODO]: Remember group names in GroupNames


    ActiveWindow.Selection.SlideRange.Shapes(TargetName).Select
    Set Sel = Application.ActiveWindow.Selection
    
    ' This function expects to receive a grouped ShapeRange
    ' We ungroup to reveal the structure at the layer below
    Sel.ShapeRange.Ungroup
    ActiveWindow.Selection.SlideRange.Shapes(TargetName).Select
           
    If Sel.ShapeRange.Type = msoGroup Then
        ' We need to go further down, the element being edited is still within a group
        ' Get the name of the Target group in which it is
        TargetGroupName = Sel.ShapeRange(1).Name
        
        Dim ShapeNames_In As New Collection  ' shapes in the same group
        Dim ShapeNames_Out As New Collection ' shapes not in the same group
        
        ' Split range according to whether elements are in the same group or not
        j_in = 0
        j_out = 0
        For Each n In ShapeNames
            ActiveWindow.Selection.SlideRange.Shapes(n).Select
            If Sel.ShapeRange.Type = msoGroup Then
                ' object is in group
                If Sel.ShapeRange(1).Name = TargetGroupName Then
                    ShapeNames_In.Add n
                Else
                    ShapeNames_Out.Add n
                End If
            Else
                ' object not in group, so it can't be in the same group as Target
                ShapeNames_Out.Add n
            End If
        Next
        
        ' Build shape range with all elements in that group, go one level down
        Dim Layers_in As New Collection
        Dim SelectionNames_in As New Collection
        
        Tmp = RecordGroupHierarchy_and_Ungroup_Fast(ShapeNames_In, TargetName, Layers_in, SelectionNames_in)
        RecordGroupHierarchy_and_Ungroup_Fast = Tmp + 1
    
        ' For all elements in that group, tag them
        For Each n In ShapeNames_In
            Layers.Add Layers_in.Item(n), n
            SelectionNames.Add SelectionNames_in.Item(n), n
        Next
        
        ' For all elements not in that group, tag them
        For Each n In ShapeNames_Out
            ActiveWindow.Selection.SlideRange.Shapes(n).Select
            Layers.Add RecordGroupHierarchy_and_Ungroup_Fast, n
            If Sel.ShapeRange.Type = msoGroup Then
                SelectionNames.Add Sel.ShapeRange(1).Name, n
            Else
                SelectionNames.Add n, n
            End If
        Next
        
    Else ' we reached the final layer: the element being edited is by itself,
         ' all other elements will need to be handled either through their group
         ' name if in a group, or their name if not
        RecordGroupHierarchy_and_Ungroup_Fast = 1
        For Each n In ShapeNames
            Layers.Add RecordGroupHierarchy_and_Ungroup_Fast, n
            SelectionNames.Add n, n
        Next
    End If

End Function


Sub MatchZOrder(oldShape As Shape, newShape As Shape)
    ' Make the Z order of newShape equal to 1 higher than that of oldShape
    
    newShape.ZOrder msoBringToFront
    While (newShape.ZOrderPosition > oldShape.ZOrderPosition + 1)
        newShape.ZOrder msoSendBackward
    Wend
End Sub

'Not used
Sub TransferGroupFormat(oldShape As Shape, newShape As Shape)
    On Error Resume Next
    ' Transfer group formatting
    If oldShape.Glow.Radius > 0 Then
        newShape.Glow.Color = oldShape.Glow.Color
        newShape.Glow.Radius = oldShape.Glow.Radius
        newShape.Glow.Transparency = oldShape.Glow.Transparency
    End If
    If oldShape.Reflection.Type <> msoReflectionTypeNone Then
        newShape.Reflection.Blur = oldShape.Reflection.Blur
        newShape.Reflection.Offset = oldShape.Reflection.Offset
        newShape.Reflection.Size = oldShape.Reflection.Size
        newShape.Reflection.Transparency = oldShape.Reflection.Transparency
        newShape.Reflection.Type = oldShape.Reflection.Type
    End If
    
    If oldShape.SoftEdge.Type <> msoSoftEdgeTypeNone Then
        newShape.SoftEdge.Radius = oldShape.SoftEdge.Radius
    End If
    
    If oldShape.Shadow.Visible Then
        newShape.Shadow.Visible = oldShape.Shadow.Visible
        newShape.Shadow.Blur = oldShape.Shadow.Blur
        newShape.Shadow.ForeColor = oldShape.Shadow.ForeColor
        newShape.Shadow.OffsetX = oldShape.Shadow.OffsetX
        newShape.Shadow.OffsetY = oldShape.Shadow.OffsetY
        newShape.Shadow.RotateWithShape = oldShape.Shadow.RotateWithShape
        newShape.Shadow.Size = oldShape.Shadow.Size
        newShape.Shadow.Style = oldShape.Shadow.Style
        newShape.Shadow.Transparency = oldShape.Shadow.Transparency
        newShape.Shadow.Type = oldShape.Shadow.Type
    End If
    
    If oldShape.ThreeD.Visible Then
        'newShape.ThreeD.BevelBottomDepth = oldShape.ThreeD.BevelBottomDepth
        'newShape.ThreeD.BevelBottomInset = oldShape.ThreeD.BevelBottomInset
        'newShape.ThreeD.BevelBottomType = oldShape.ThreeD.BevelBottomType
        'newShape.ThreeD.BevelTopDepth = oldShape.ThreeD.BevelTopDepth
        'newShape.ThreeD.BevelTopInset = oldShape.ThreeD.BevelTopInset
        'newShape.ThreeD.BevelTopType = oldShape.ThreeD.BevelTopType
        'newShape.ThreeD.ContourColor = oldShape.ThreeD.ContourColor
        'newShape.ThreeD.ContourWidth = oldShape.ThreeD.ContourWidth
        'newShape.ThreeD.Depth = oldShape.ThreeD.Depth
        'newShape.ThreeD.ExtrusionColor = oldShape.ThreeD.ExtrusionColor
        'newShape.ThreeD.ExtrusionColorType = oldShape.ThreeD.ExtrusionColorType
        newShape.ThreeD.Visible = oldShape.ThreeD.Visible
        newShape.ThreeD.Perspective = oldShape.ThreeD.Perspective
        newShape.ThreeD.FieldOfView = oldShape.ThreeD.FieldOfView
        newShape.ThreeD.LightAngle = oldShape.ThreeD.LightAngle
        'newShape.ThreeD.ProjectText = oldShape.ThreeD.ProjectText
        'If oldShape.ThreeD.PresetExtrusionDirection <> msoPresetExtrusionDirectionMixed Then
        '    newShape.ThreeD.SetExtrusionDirection oldShape.ThreeD.PresetExtrusionDirection
        'End If
        newShape.ThreeD.PresetLighting = oldShape.ThreeD.PresetLighting
        If oldShape.ThreeD.PresetLightingDirection <> msoPresetLightingDirectionMixed Then
            newShape.ThreeD.PresetLightingDirection = oldShape.ThreeD.PresetLightingDirection
        End If
        If oldShape.ThreeD.PresetLightingSoftness <> msoPresetLightingSoftnessMixed Then
            newShape.ThreeD.PresetLightingSoftness = oldShape.ThreeD.PresetLightingSoftness
        End If
        If oldShape.ThreeD.PresetMaterial <> msoPresetMaterialMixed Then
            newShape.ThreeD.PresetMaterial = oldShape.ThreeD.PresetMaterial
        End If
        If oldShape.ThreeD.PresetCamera <> msoPresetCameraMixed Then
            newShape.ThreeD.SetPresetCamera oldShape.ThreeD.PresetCamera
        End If
        newShape.ThreeD.RotationX = oldShape.ThreeD.RotationX
        newShape.ThreeD.RotationY = oldShape.ThreeD.RotationY
        newShape.ThreeD.RotationZ = oldShape.ThreeD.RotationZ
        'newShape.ThreeD.Z = oldShape.ThreeD.Z
    End If
End Sub

