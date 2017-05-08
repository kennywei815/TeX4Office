Attribute VB_Name = "Main"
'''''==============================================================================================================================================
'''''                                            Platform-dependent Macros
'''''==============================================================================================================================================
#Const PowerPoint = 0
#Const Word = 1
#Const Excel = 2

#Const PLATFORM = PowerPoint


#Const TeX4Office = 0
#Const ImportImage = 1

#Const PROGRAM = ImportImage


'''''==============================================================================================================================================
'''''                                            Main Function
'''''==============================================================================================================================================
#If PROGRAM = TeX4Office Then
Sub EditLaTeX()

#ElseIf PROGRAM = ImportImage Then
Sub InsertPicture()

#End If
    '[DONE] �����β{�b���󤤪�TeX�{���X & �ɦW  �Ӷ}��
    '[DONE] TeX4Office Editor�i�H�磌��W & �s�a���ܹ��ɦW�r
    '[DONE] �ˬdselection�O���O��@��LaTeX shape�A�Ӥ��Ogroup
    
    'TODO: �� Word �� RecordGroupHierarchy_and_Ungroup �\��A�������u��
    
    'TODO: PowerPoint�BWord�BExcel�U��������@�@�� Build() ��ơA�i�H�۰ʸ��J�̷s��.bas�ɽsĶ => �H��i�H�u��@��.bas��
    
    'TODO: �]�n�w���ɡB�g�w�ˤ�U�B�ϥΤ�U
    'TODO: �}����������:
    '       1. Windows XP + Office 2003
    '       2. Windows 7  + Office 2007
    '       3. Windows 7  + Office 2010
    '       4. Windows 7  + Office 2013
    '       5. Windows 10 + Office 2016
    '       6. Windows 10 + Office 365
    '       7. ��ī�G     + Office 2011 for Mac
    '       8. ��ī�G     + Office 2016 for Mac
    '       9. ��ī�G     + Office 2008 for Mac
        
    'TODO: Word�BExcel���ӧ�Ϥ���b�@�}�l����Ъ���
    'TODO: �n ungroup �M regroup �����ӭn��ܤ@�Ӽu������A��ܥثe���A�F�p�GTime out�Φ����~�A�]�n�i�D�ϥΪ�
    'TODO: ��n�����~�B�z
    
    
    
    'TODO: ��ڤƥ\��
    
    
    'TODO: read PointSize
    'TODO: read DPI
    
    'TODO: VBA��Editor
    'TODO: ��Theme???�������n�ݤ@�I
    
    'TODO: [Konwn Issue] Excel�������D�p�󮳥XOldGroup�BOldGroup.GroupItems�A�ȮɵL�kungroup�Aregroup�A�u�ઽ���R���ª��A���s��
    
    'Insert png file & do the neccessary post-processing
#If PLATFORM = PowerPoint Then
    Set sld = Application.ActiveWindow.View.Slide
    Set osld = ActiveWindow.Selection.SlideRange(1)
    Set AllShapes = ActiveWindow.Selection.SlideRange.Shapes
            
#ElseIf PLATFORM = Word Then
    Set sld = ActiveDocument
    Set osld = ActiveDocument
    Set AllShapes = ActiveDocument.Shapes
            
#ElseIf PLATFORM = Excel Then
    Set sld = ActiveSheet
    Set osld = ActiveSheet
    Set AllShapes = ActiveSheet.Shapes
            
#End If
        
    ' If we are in Edit mode, store parameters of old image
    Dim PosX As Single
    Dim PosY As Single
    Set Sel = ActiveWindow.Selection
    
    Dim oldShape As Shape
    
    Dim s As Shape
    IsInGroup = False
    
    
    '==============================================================================================================================================
    ' Get information from old shape (if existed)

    If selectionIsLaTeXShape() Then
        ' Edit old shape
        
        If selectionIsGroup() Then
        
            ' Old image is part of a group
#If PLATFORM = Excel Then
            Set oldShape = Sel.ShapeRange(1) 'For Excel
#Else
            Set oldShape = Sel.ChildShapeRange(1) 'For PowerPoint & Word
            
#End If
            
            IsInGroup = True
            Dim ShapeNames As New Collection ' gather all shapes to be regrouped later on
            
            ' Store the group's animation and Zorder info in a dummy object tmpGroup
            Dim oldGroup As Shape
            Dim tmpGroup As Shape
            
            Set oldGroup = Sel.ShapeRange(1) 'TODO_NOW: Excel�����P
            Set tmpGroup = AllShapes.AddShape(msoShapeDiamond, 1, 1, 1, 1)
            
            MoveAnimation oldGroup, tmpGroup
            MatchZOrder oldGroup, tmpGroup
        
            ' Tag all elements in the group with their hierarchy level and their name or group name
            Dim MaxGroupLevel As Long
            Dim Layers As New Collection
            Dim oldShapeSelectionName As String
            Dim GroupNames As New Collection
            Dim SelectionNames As New Collection
            
            oldShapeSelectionName = InvalidName
            
'[TODO][FIXME]: PowerPoint����Sel.ShapeRange�|���AWord����Sel.ShapeRange.GroupItems  =>  ��X�����I�BOffice�T�j���~Seletion�����BSeletion���Shape�BShapes�BShapeRange�BGroupShapes���޿�A�O���U�ӡA���}�ѫ�H�Ѧ�
#If PLATFORM = PowerPoint Then
            For Each s In Sel.ShapeRange.GroupItems  'TODO_NOW: �o�O��PowerPoint���Ϊ� => �[#IF
                                                     'TODO_NOW: Excel�u���s�դ~��s���o�Ӧ���
                If s.Name <> oldShape.Name Then
                    ShapeNames.Add s.Name
                End If
            Next
            
            ''====================================================================================== BEGIN DEBUG ======================================================================================
            'MsgBox "Constructing ShapeNames" 'DEBUG
            'For Each n In ShapeNames 'DEBUG
            '    MsgBox "ShapeNames: " & n  'DEBUG
            'Next 'DEBUG
            ''====================================================================================== END DEBUG ======================================================================================
            
            'MaxGroupLevel = RecordGroupHierarchy_and_Ungroup_Fast(oldGroup, oldShape.Name, oldShapeSelectionName, ShapeNames, Layers, GroupNames)
            MaxGroupLevel = RecordGroupHierarchy_and_Ungroup_Fast(ShapeNames, oldShape.Name, Layers, SelectionNames)
            
#ElseIf PLATFORM = Word Then
            MaxGroupLevel = RecordGroupHierarchy_and_Ungroup(oldGroup, oldShape.Name, oldShapeSelectionName, ShapeNames, Layers, GroupNames)
#End If
            
            oldShape.Select
            
        Else
            Set oldShape = Sel.ShapeRange(1)
        End If
        PosX = oldShape.Left
        PosY = oldShape.Top
    Else
        PosX = 200
        PosY = 200
    End If
    
    
    '==============================================================================================================================================
    ' Run LaTeX Editor or other PNG generators
    'TODO: [FIXME] ��Word�BPowerPoint�BExcel�ɦW�ɡA����|�o�Ϳ��~
    Dim TempDir As String
    Dim FilePrefix As String
    Dim code As String
    
    Code2PNG oldShape, TempDir, FilePrefix, code
    
    sourceFileName = FilePrefix & ".tex"
    pictureFileName = FilePrefix & ".png"
    
    
    If Dir(TempDir & pictureFileName) = Empty Then
        Exit Sub
    End If
    
    
    '==============================================================================================================================================
    ' Get scaling factors
    
    default_screen_dpi = 96
    OutputDpi = 600    '[TODO]: read OutputDpi from TeX4Office Editor's config.json
    OldDpi = OutputDpi '[TODO]: read OldDpi    from shape.AlternativeText ???
    
    MagicScalingFactorPNG = default_screen_dpi / OutputDpi
    
    
    '==============================================================================================================================================
    ' Insert image and rescale it
    Dim newShape As Shape
    
    Set newShape = AddDisplayShape(TempDir + pictureFileName, PosX, PosY)
    
    'Delete temporary files
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    If Dir(TempDir & FilePrefix & ".*") <> Empty Then
        fs.DeleteFile TempDir & FilePrefix & ".*"
    End If
    
    ' Resize to the true size of the png file and adjust using the manual scaling factors set in Main Settings
    With newShape
        .ScaleHeight 1#, msoTrue
        .ScaleWidth 1#, msoTrue
        
        .LockAspectRatio = msoFalse
        
        ' Apply scaling factors
        If selectionIsLaTeXShape() Then
            '[TODO]: read current & old DPI
            .ScaleHeight oldShape.Height / .Height * OldDpi / OutputDpi, msoTrue
            .ScaleWidth oldShape.Width / .Width * OldDpi / OutputDpi, msoTrue
        Else
            '[TODO]: read DPI       from TeX4Office Editor's config.json
            '[TODO]: read PointSize from TeX4Office Editor's config.json
            
#If PROGRAM = TeX4Office Then
            PointSize = 10
            tScale = PointSize / 10 * MagicScalingFactorPNG  ' 1/10 is for the default LaTeX point size (10 pt)
            
            .ScaleHeight tScale, msoTrue
            .ScaleWidth tScale, msoTrue
#End If
            
        End If
        
        .LockAspectRatio = msoTrue
    End With
        
    ' We are editing+resetting size of an old display, we keep rotation
    If selectionIsLaTeXShape() Then
        newShape.Rotation = oldShape.Rotation
        newShape.LockAspectRatio = oldShape.LockAspectRatio ' Unlock aspect ratio if old display had it unlocked
    End If
    
    
    '==============================================================================================================================================
    ' Add tags
    Call AddTagsToShape(newShape, code, FilePrefix)
    
    '==============================================================================================================================================
    ' Copy animation settings and formatting from old image, then delete it
    If selectionIsLaTeXShape() Then

        If IsInGroup Then

            ' Transfer format to new shape
            Dim GroupName As String
            
            GroupName = oldShapeSelectionName
            
            MatchZOrder oldShape, newShape
            
            oldShape.PickUp
            newShape.Apply
            oldShape.Delete
            
            Dim newGroup As Shape

            ' Group all non-modified elements from old group, plus modified element
            ShapeNames.Add newShape.Name
            
            Layers.Add 1, newShape.Name
            GroupNames.Add GroupName, newShape.Name
            SelectionNames.Add newShape.Name, newShape.Name
                        
            ' Hierarchically re-group elements
            For Level = 1 To MaxGroupLevel
            
'                '====================================================================================== BEGIN DEBUG ======================================================================================
'                MsgBox "Begin regrouping" 'DEBUG
'                MsgBox "MaxGroupLevel=" & MaxGroupLevel
'                MsgBox "ShapeNames.Count=" & ShapeNames.Count
'                MsgBox "Layers.Count=" & Layers.Count
'                MsgBox "SelectionNames.Count=" & SelectionNames.Count
'                MsgBox "GroupNames.Count=" & GroupNames.Count
'                For Each n In ShapeNames 'DEBUG
'#If PLATFORM <> Word Then
'                    MsgBox "ShapeNames: (Name, Layer, SelectionName) = " & n & ", " & Layers(n) & ", " & SelectionNames(n) 'DEBUG
'#Else
'                    MsgBox "ShapeNames: (Name, Layer, GroupName) = " & n & ", " & Layers(n) & ", " & GroupNames(n) 'DEBUG
'#End If
'                Next 'DEBUG
'                '====================================================================================== END DEBUG ======================================================================================
            
                Dim CurrentLevelShapeNames As New Collection
                ClearCollection CurrentLevelShapeNames
                
                For Each n In ShapeNames
                    Dim ShapeLevel As Integer
                    ShapeLevel = 0
                    
                    On Error Resume Next
                    ShapeLevel = Layers.Item(n)
                    
                    Dim n_ToBeGroup As String
                    
#If PLATFORM <> Word Then
                    n_ToBeGroup = SelectionNames(n)
#Else
                    n_ToBeGroup = n
#End If
                    
                    If ShapeLevel = Level Then
                        If CurrentLevelShapeNames.Count > 0 Then
                            If Not IsInCollection(CurrentLevelShapeNames, n_ToBeGroup) Then
                                CurrentLevelShapeNames.Add n_ToBeGroup
                            End If
                        Else
                            CurrentLevelShapeNames.Add n_ToBeGroup
                        End If
                    End If
                Next
                
'                '====================================================================================== BEGIN DEBUG ======================================================================================
'                MsgBox "CurrentLevelShapeNames.Count=" & CurrentLevelShapeNames.Count 'DEBUG
'                For Each n In CurrentLevelShapeNames 'DEBUG
'#If PLATFORM <> Word Then
'                    MsgBox "CurrentLevelShapeNames: ShapeNames: (Name, Layer, SelectionName) = " & n & ", " & Layers(n) & ", " & SelectionNames(n) 'DEBUG
'#Else
'                    MsgBox "CurrentLevelShapeNames: ShapeNames: (Name, Layer, GroupName) = " & n & ", " & Layers(n) & ", " & GroupNames(n) 'DEBUG
'#End If
'                Next 'DEBUG
'                '====================================================================================== END DEBUG ======================================================================================
                
                If CurrentLevelShapeNames.Count > 1 Then
                    Set newGroup = sld.Shapes.Range(ToArray(CurrentLevelShapeNames)).Group()
#If PLATFORM = Word Then
                    newGroup.Name = GroupNames(CurrentLevelShapeNames(1))
#End If
                    ShapeNames.Add newGroup.Name
                    
                    Layers.Add Level + 1, newGroup.Name
                    GroupNames.Add newGroup.Name, newGroup.Name
                    SelectionNames.Add newGroup.Name, newGroup.Name
                    
                    For Each n_CurrentLevel In CurrentLevelShapeNames
                        ShapeNames.Remove n_CurrentLevel
                    Next
                End If
                
            Next
            
            ' Use temporary group to retrieve the group's original animation and Zorder
            MoveAnimation tmpGroup, newGroup
            MatchZOrder tmpGroup, newGroup
            tmpGroup.Delete
        Else
            MoveAnimation oldShape, newShape
            MatchZOrder oldShape, newShape
            
            oldShape.PickUp
            newShape.Apply
            oldShape.Delete
        End If
    End If
End Sub

