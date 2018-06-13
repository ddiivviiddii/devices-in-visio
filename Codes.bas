Attribute VB_Name = "Codes"
Public Sub macrosik()
Call DrawDevice(0, 0)
End Sub
Sub DrawDevice(XBox As Double, YBox As Double)
    Dim DiagramServices As Integer
    DiagramServices = ActiveDocument.DiagramServicesEnabled
    ActiveDocument.DiagramServicesEnabled = visServiceVersion140 + visServiceVersion150
    Dim UndoScopeID2 As Long
    UndoScopeID2 = Application.BeginUndoScope("Add device")
    
    Dim RectShp As Visio.Shape
    Dim i As Integer                'var for loops
    Dim strin As String             'universal var for make text for connector's name and other purpose
    Dim intRowIndex1 As Integer     'technical var for adding row in the properties of the box
    Dim vsoRow1 As Visio.Row        'technical var for adding row in the properties of the box
    Dim LineShp As Visio.Shape
    Dim TxtShp  As Visio.Shape      'text object for labels
    Dim Chars   As Visio.Characters 'char object for labels
    Const XBoxW = 25                'Default width of the box in mm
    
    '5 mm in inches as Visio works in inches only
    Dim Fivemm As Double
    Fivemm = MiMi(5)
    
    Dim LConnects As Integer    'Number of connectors from left side
    Dim RConnects As Integer    'Number of connectors from right side
    Dim YBoxH As Double         'Height of the box
    Dim XBox2 As Double         'X of another point of the box
    Dim YBox2 As Double         'y of another point of the box
    Dim j As Integer            'Counter for all shapes in the selection
    Dim y As Integer            'Stores number of all shapes
    Dim InTxt As String         'Name of input connectors
    Dim OutTxt As String        'Name of output connectors
    Dim DevTxt As String        'Position of device like DIST.01
    Dim ModelTxt As String      'Name of device like GDR216
    
    'Form with initial settings
    ParametersForm.Show
    If IsNumeric(ParametersForm.txt_in.Text) Then
        LConnects = CInt(ParametersForm.txt_in.Text)
    End If
    If IsNumeric(ParametersForm.txt_out.Text) Then
        RConnects = CInt(ParametersForm.txt_out.Text)
    End If
    DevTxt = ParametersForm.txt_position.Text
    ModelTxt = ParametersForm.txt_model.Text
    InTxt = ParametersForm.txt_in_name.Text
    OutTxt = ParametersForm.txt_out_name.Text
    Unload ParametersForm
    
    'Arrays for input's names
    Dim Inn() As String                 'Array for name of inputs
    ReDim Inn(LConnects)
    Dim InnCon() As String              'Array for type of inputs
    ReDim InnCon(LConnects)
    'Array for output's names
    Dim Outn() As String                'Array for name of outputs
    ReDim Outn(RConnects)
    Dim OutCon() As String
    ReDim OutCon(RConnects)             'Array for type of outputs
    Dim InTxtBox() As MSForms.TextBox   'Array for current input's name in the form
    ReDim InTxtBox(LConnects)
    Dim InConTxtBox() As MSForms.TextBox 'Arrat for current type of inputs in the form
    ReDim InConTxtBox(LConnects)
    Dim OutTxtBox() As MSForms.TextBox   'Array for current output's name in the form
    ReDim OutTxtBox(RConnects)
    Dim OutConTxtBox() As MSForms.TextBox 'Arrat for current type of outputs in the form
    ReDim OutConTxtBox(RConnects)
    
    Dim TxtPosition As Integer          'Position of text in from with names of connections
    
    'Form with connections settings
    Load DeviceForm
    'Set dimensions of the form
    If LConnects > RConnects Then DeviceForm.Height = 20.25 + LConnects * 20 Else DeviceForm.Height = 20.25 + RConnects * 20
    
    'Fill the form
    TxtPosition = 0
    For i = 1 To LConnects
        If i < 10 Then strin = "IN" & " " & "0" & CStr(i) Else strin = "IN" & " " & CStr(i)
        Set InTxtBox(i - 1) = DeviceForm.Controls.Add("Forms.TextBox.1", "txt_in" & (i))
        With InTxtBox(i - 1)
            .Left = 10
            .Top = TxtPosition
            .Width = 50
            .Text = strin
        End With
        Set InConTxtBox(i - 1) = DeviceForm.Controls.Add("Forms.TextBox.1", "txt_incon" & (i))
        With InConTxtBox(i - 1)
            .Left = 70
            .Top = TxtPosition
            .Width = 50
            .Text = "BNC"
        End With
        TxtPosition = TxtPosition + 20
    Next i
    
    TxtPosition = 0
    For i = 1 To RConnects
        If i < 10 Then strin = "OUT" & " " & "0" & CStr(i) Else strin = "OUT" & " " & CStr(i)
        Set OutTxtBox(i - 1) = DeviceForm.Controls.Add("Forms.TextBox.1", "txt_out" & (i))
        With OutTxtBox(i - 1)
            .Left = 140
            .Top = TxtPosition
            .Width = 50
            .Text = strin
        End With
        Set OutConTxtBox(i - 1) = DeviceForm.Controls.Add("Forms.TextBox.1", "txt_outcon" & (i))
        With OutConTxtBox(i - 1)
            .Left = 200
            .Top = TxtPosition
            .Width = 50
            .Text = "BNC"
        End With
        TxtPosition = TxtPosition + 20
    Next i
    
    DeviceForm.Show
    
    'Move text data from form to arrays
    For i = 1 To LConnects
       Inn(i - 1) = InTxtBox(i - 1).Text
       InnCon(i - 1) = InConTxtBox(i - 1).Text
    Next i
    
    For i = 1 To RConnects
       Outn(i - 1) = OutTxtBox(i - 1).Text
       OutCon(i - 1) = OutConTxtBox(i - 1).Text
    Next i
    
    Unload DeviceForm
    
    'Calculation the height of the box
    If LConnects > RConnects Then
        YBoxH = 5 + 5 * LConnects
    Else
        YBoxH = 5 + 5 * RConnects
    End If
    
    XBox2 = XBox + XBoxW
    YBox2 = YBox - YBoxH
    
    y = 3 + LConnects * 3 + RConnects * 3   'Totaly one box and LConnects + RConnects lines + all labels. Started from one.
    j = 0                                   'Pointer for current object. First is zero
    
    'Convert all coordinates from mm to inches as we remember about this fucking feature of Visio
    XBox = MiMi(XBox)
    YBox = MiMi(YBox)
    XBox2 = MiMi(XBox2)
    YBox2 = MiMi(YBox2)
    
    'Draw box and get its ID
    Set RectShp = Application.ActiveWindow.Page.DrawRectangle(XBox, YBox, XBox2, YBox2)
    
    Dim sGUID() As String       'This is array for all shape's ID's in the device
    ReDim sGUID(y)
    
    'Put the box ID to array for future grouping
    sGUID(j) = RectShp.UniqueID(visGetOrMakeGUID)
    
    'Make it black and nice
    GoToBlack (sGUID(j))
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionObject, visRowLine, visLineWeight).FormulaU = "1 pt"
    
    'Drawing label for position of the device
    Set TxtShp = Application.ActiveWindow.Page.DrawRectangle(XBox + MiMi(XBoxW / 2) - MiMi(7), YBox + MiMi(1.5), XBox + MiMi(XBoxW / 2) + MiMi(7), YBox + MiMi(1.5))
    j = j + 1
    TxtShp.TextStyle = "Normal"
    TxtShp.LineStyle = "Text Only"
    TxtShp.FillStyle = "Text Only"
    sGUID(j) = TxtShp.UniqueID(visGetOrMakeGUID)    'Add it to array for future grouping
    Set Chars = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).Characters
    Chars.Begin = 0
    Chars.End = 3
    Chars.Text = DevTxt
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "8 pt"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,0))"
    GoToHide (sGUID(j))
    
    'Drawing label for model of the device
    Set TxtShp = Application.ActiveWindow.Page.DrawRectangle(XBox + MiMi(XBoxW / 2) - MiMi(7), YBox - MiMi(2), XBox + MiMi(XBoxW / 2) + MiMi(7), YBox - MiMi(2))
    j = j + 1
    TxtShp.TextStyle = "Normal"
    TxtShp.LineStyle = "Text Only"
    TxtShp.FillStyle = "Text Only"
    sGUID(j) = TxtShp.UniqueID(visGetOrMakeGUID)    'Add it to array for future grouping
    Set Chars = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).Characters
    Chars.Begin = 0
    Chars.End = 3
    Chars.Text = ModelTxt
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "8 pt"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,0))"
    GoToHide (sGUID(j))
    
    'Creating connection on the right side of the device
    If RConnects > 0 Then
            For i = 1 To RConnects
                Set LineShp = Application.ActiveWindow.Page.DrawLine(XBox2, YBox - i * Fivemm, XBox2 + Fivemm, YBox - i * Fivemm)
                'Increment of pointer of current object
                j = j + 1
                'Convert the box ID to UniqueID for future reference
                sGUID(j) = LineShp.UniqueID(visGetOrMakeGUID)
                'Adding row for new connection point
                intRowIndex1 = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).AddRow(visSectionConnectionPts, visRowLast, visTagCnnctPt)
                Set vsoRow1 = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).Section(visSectionConnectionPts).Row(intRowIndex1)
                'Position of the connection point according line's coordinates
                vsoRow1.Cell(visCnnctY).FormulaU = "0 mm"
                vsoRow1.Cell(visCnnctX).FormulaU = "Width*1"
                vsoRow1.Cell(visCnnctDirX).FormulaU = -1#
                vsoRow1.Cell(visCnnctDirY).FormulaU = 0#
                vsoRow1.Cell(visCnnctType).FormulaU = visCnnctTypeInward
                GoToBlack (sGUID(j))
                j = j + 1
                Set TxtShp = Application.ActiveWindow.Page.DrawRectangle(XBox2 - 2 * Fivemm, YBox - i * Fivemm, XBox2, YBox - i * Fivemm)
                TxtShp.TextStyle = "Normal"
                TxtShp.LineStyle = "Text Only"
                TxtShp.FillStyle = "Text Only"
                sGUID(j) = TxtShp.UniqueID(visGetOrMakeGUID)
                Set Chars = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).Characters
                Chars.Begin = 0
                Chars.End = 3
                Chars.Text = Outn(i - 1)
                Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "6 pt"
                Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,0))"
                GoToHide (sGUID(j))
                'Adding label for type of output
                j = j + 1
                Set TxtShp = Application.ActiveWindow.Page.DrawRectangle(XBox2 - MiMi(5 / 4), YBox - i * Fivemm + MiMi(5 / 10), XBox2 - MiMi(5 / 4) + MiMi(6), YBox - i * Fivemm + MiMi(5 / 10))
                TxtShp.TextStyle = "Normal"
                TxtShp.LineStyle = "Text Only"
                TxtShp.FillStyle = "Text Only"
                sGUID(j) = TxtShp.UniqueID(visGetOrMakeGUID)
                Set Chars = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).Characters
                Chars.Begin = 0
                Chars.End = 3
                Chars.Text = OutCon(i - 1)
                Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "3 pt"
                Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,0))"
                GoToHide (sGUID(j))
            Next i
    End If
    'Creating connection on the left side of the box
    If LConnects > 0 Then
            For i = 1 To LConnects
                Set LineShp = Application.ActiveWindow.Page.DrawLine(XBox - Fivemm, YBox - i * Fivemm, XBox, YBox - i * Fivemm)
                'Convert the box ID to UniqueID for future reference
                j = j + 1
                sGUID(j) = LineShp.UniqueID(visGetOrMakeGUID)
                'Adding row for new connection point
                intRowIndex1 = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).AddRow(visSectionConnectionPts, visRowLast, visTagCnnctPt)
                Set vsoRow1 = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).Section(visSectionConnectionPts).Row(intRowIndex1)
                vsoRow1.Cell(visCnnctY).FormulaU = "0 mm"
                vsoRow1.Cell(visCnnctX).FormulaU = "0 mm"
                vsoRow1.Cell(visCnnctDirX).FormulaU = 1#
                vsoRow1.Cell(visCnnctDirY).FormulaU = 0#
                vsoRow1.Cell(visCnnctType).FormulaU = visCnnctTypeInward
                GoToBlack (sGUID(j))
                'Adding label for input
                j = j + 1
                Set TxtShp = Application.ActiveWindow.Page.DrawRectangle(XBox + 2 * Fivemm, YBox - i * Fivemm, XBox, YBox - i * Fivemm)
                TxtShp.TextStyle = "Normal"
                TxtShp.LineStyle = "Text Only"
                TxtShp.FillStyle = "Text Only"
                sGUID(j) = TxtShp.UniqueID(visGetOrMakeGUID)
                Set Chars = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).Characters
                Chars.Begin = 0
                Chars.End = 3
                Chars.Text = Inn(i - 1)
                Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "6 pt"
                Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,0))"
                GoToHide (sGUID(j))
                'Adding label for type of input
                j = j + 1
                Set TxtShp = Application.ActiveWindow.Page.DrawRectangle(XBox - MiMi(5), YBox - i * Fivemm + MiMi(5 / 10), XBox + MiMi(5 / 4), YBox - i * Fivemm + MiMi(5 / 10))
                TxtShp.TextStyle = "Normal"
                TxtShp.LineStyle = "Text Only"
                TxtShp.FillStyle = "Text Only"
                sGUID(j) = TxtShp.UniqueID(visGetOrMakeGUID)
                Set Chars = Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).Characters
                Chars.Begin = 0
                Chars.End = 3
                Chars.Text = InnCon(i - 1)
                Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterSize).FormulaU = "3 pt"
                Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(j)).CellsSRC(visSectionCharacter, 0, visCharacterColor).FormulaU = "THEMEGUARD(RGB(0,0,0))"
                GoToHide (sGUID(j))
            Next i
    End If
  
    ActiveWindow.DeselectAll
    For i = 0 To y - 1
          ActiveWindow.Select Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(sGUID(i)), visSelect
    Next i
    ActiveWindow.Selection.Group
    Application.EndUndoScope UndoScopeID2, True
    ActiveDocument.DiagramServicesEnabled = DiagramServices
End Sub

Function MiMi(mm As Double)
'This function converts mm to inches
MiMi = Application.ConvertResult(mm, "mm", "in")
End Function

Sub GoToBlack(id As String)
'This function paints shape to black
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillPattern).FormulaU = "0"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowGradientProperties, visFillGradientEnabled).FormulaU = "FALSE"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowGradientProperties, visRotateGradientWithShape).FormulaU = "FALSE"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowGradientProperties, visUseGroupGradient).FormulaU = "FALSE"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowLine, visLineColor).FormulaU = "THEMEGUARD(RGB(0,0,0))"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowGradientProperties, visLineGradientEnabled).FormulaU = "FALSE"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillShdwForegnd).FormulaU = "THEMEGUARD(RGB(0,0,0))"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillShdwPattern).FormulaU = "0"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillShdwForegndTrans).FormulaU = "0%"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillShdwType).FormulaU = "1"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillShdwOffsetX).FormulaU = "0 pt"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillShdwOffsetY).FormulaU = "0 pt"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillShdwScaleFactor).FormulaU = "0%"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowFill, visFillShdwBlur).FormulaU = "0 pt"
End Sub

Sub GoToHide(id As String)
'This function hides border of the shape
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowLine, visLinePattern).FormulaU = "0"
    Application.ActiveWindow.Page.Shapes.ItemFromUniqueID(id).CellsSRC(visSectionObject, visRowGradientProperties, visLineGradientEnabled).FormulaU = "FALSE"
End Sub



