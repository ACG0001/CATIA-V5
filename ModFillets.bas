Attribute VB_Name = "ModFillets"
'================================================================
' MACRO : FindNonStandardFillets
' VERSION: 10 - With UserForm
'
' SETUP INSTRUCTIONS:
'   1. CATIA > Tools > Macros > Visual Basic Editor (Alt+F11)
'
'   --- MODULE ---
'   2. Insert > Module
'   3. Name it: ModFillets
'   4. Paste the MODULE CODE section into it
'
'   --- USERFORM ---
'   5. Insert > UserForm
'   6. Name it: FrmFilletReport
'   7. Add these controls with EXACT names:
'
'      Control        Name             Properties
'      -----------------------------------------------------------
'      Label          lblTitle         Font Bold=True, Size=10
'      ListBox        lstFillets       Width=460, Height=280
'                                      MultiSelect=0 (single)
'      Label          lblGrandTotal    Font Bold=True
'      CommandButton  btnCopy          Caption="Copy to clipboard"
'      CommandButton  btnExport        Caption="Export to .txt"
'      CommandButton  btnClose         Caption="Close"
'
'   8. Paste the USERFORM CODE section into the UserForm code window
'      (double-click the form to open its code window)
'
'   9. Open a CATPart and run: FindNonStandardFillets
'================================================================


'################################################################
' MODULE CODE — paste into ModFillets
'################################################################

Option Explicit

'----------------------------------------------------------------
' Global variables shared with the UserForm
'----------------------------------------------------------------
Public gDoc        As Document
Public gPart       As Part
Public gNames()    As String
Public gRadii()    As Double
Public gShapes()   As AnyObject
Public gColorIdx() As Integer
Public gCount      As Long

Public gColorNames(7) As String
Public gColors(23)    As Integer

'================================================================
' ENTRY POINT — run this Sub
'================================================================
Sub FindNonStandardFillets()

    '------------------------------------------------------------
    ' VARIABLE DECLARATIONS
    '------------------------------------------------------------
    Dim oBodies   As Bodies
    Dim oBody     As Body
    Dim oShapes   As Shapes
    Dim oShape    As AnyObject
    Dim oVP       As VisPropertySet
    Dim iBody     As Integer
    Dim iShape    As Integer
    Dim sType     As String
    Dim dRadius   As Double
    Dim dRnd      As Double
    Dim k         As Long
    Dim l         As Long
    Dim m         As Long
    Dim tempName  As String
    Dim tempRad   As Double
    Dim tempIdx   As Integer
    Dim tempShape As AnyObject
    Dim curRad    As Double
    Dim colorIdx  As Integer

    '------------------------------------------------------------
    ' STEP 1 - Validate active document is a CATPart
    '------------------------------------------------------------
    If CATIA.Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "Macro"
        Exit Sub
    End If

    Set gDoc = CATIA.ActiveDocument

    If InStr(LCase(gDoc.Name), ".catpart") = 0 Then
        MsgBox "Please open a CATPart first.", vbExclamation, "Macro"
        Exit Sub
    End If

    Set gPart = gDoc.Part

    '------------------------------------------------------------
    ' STEP 2 - Initialise colour palette
    '------------------------------------------------------------
    gColorNames(0) = "Red"
    gColorNames(1) = "Green"
    gColorNames(2) = "Blue"
    gColorNames(3) = "Orange"
    gColorNames(4) = "Purple"
    gColorNames(5) = "Cyan"
    gColorNames(6) = "Pink"
    gColorNames(7) = "White"

    gColors(0) = 255: gColors(1) = 0: gColors(2) = 0
    gColors(3) = 0: gColors(4) = 180: gColors(5) = 0
    gColors(6) = 0: gColors(7) = 100: gColors(8) = 255
    gColors(9) = 255: gColors(10) = 140: gColors(11) = 0
    gColors(12) = 160: gColors(13) = 0: gColors(14) = 200
    gColors(15) = 0: gColors(16) = 200: gColors(17) = 200
    gColors(18) = 255: gColors(19) = 0: gColors(20) = 180
    gColors(21) = 255: gColors(22) = 255: gColors(23) = 255

    '------------------------------------------------------------
    ' STEP 3 - Clear selection and initialise result arrays
    '------------------------------------------------------------
    gDoc.Selection.Clear
    gCount = 0
    ReDim gNames(0)
    ReDim gRadii(0)
    ReDim gShapes(0)
    ReDim gColorIdx(0)

    '------------------------------------------------------------
    ' STEP 4 - Loop every Body > every Shape
    '------------------------------------------------------------
    Set oBodies = gPart.Bodies

    For iBody = 1 To oBodies.Count

        Set oBody = oBodies.Item(iBody)
        Set oShapes = oBody.Shapes

        For iShape = 1 To oShapes.Count

            Set oShape = oShapes.Item(iShape)
            sType = LCase(TypeName(oShape))

            If InStr(sType, "edgefillet") > 0 Or _
               InStr(sType, "fillet") > 0 Then

                dRadius = ReadRadius(oShape, gPart)

                If dRadius > 0 Then

                    dRnd = CDbl(Format(dRadius, "0.000"))

                    If dRnd <> 4# And dRnd <> 9# Then

                        ReDim Preserve gNames(gCount)
                        ReDim Preserve gRadii(gCount)
                        ReDim Preserve gShapes(gCount)
                        ReDim Preserve gColorIdx(gCount)
                        gNames(gCount) = oShape.Name
                        gRadii(gCount) = dRnd
                        Set gShapes(gCount) = oShape
                        gColorIdx(gCount) = 0
                        gCount = gCount + 1

                    End If
                End If
            End If

        Next iShape
    Next iBody

    '------------------------------------------------------------
    ' STEP 5 - Exit if nothing found
    '------------------------------------------------------------
    If gCount = 0 Then
        MsgBox "No non-standard fillets found." & vbCrLf & _
               "All fillets have radius 4 mm or 9 mm.", _
               vbInformation, "Result"
        Exit Sub
    End If

    '------------------------------------------------------------
    ' STEP 6 - Sort by radius ascending (bubble sort)
    '------------------------------------------------------------
    For k = 0 To gCount - 2
        For l = 0 To gCount - 2 - k
            If gRadii(l) > gRadii(l + 1) Then
                tempRad = gRadii(l)
                gRadii(l) = gRadii(l + 1)
                gRadii(l + 1) = tempRad
                tempName = gNames(l)
                gNames(l) = gNames(l + 1)
                gNames(l + 1) = tempName
                Set tempShape = gShapes(l)
                Set gShapes(l) = gShapes(l + 1)
                Set gShapes(l + 1) = tempShape
                tempIdx = gColorIdx(l)
                gColorIdx(l) = gColorIdx(l + 1)
                gColorIdx(l + 1) = tempIdx
            End If
        Next l
    Next k

    '------------------------------------------------------------
    ' STEP 7 - Assign colour per radius group and apply in CATIA
    '------------------------------------------------------------
    curRad = -1
    colorIdx = -1

    For m = 0 To gCount - 1

        If gRadii(m) <> curRad Then
            curRad = gRadii(m)
            colorIdx = colorIdx + 1
            If colorIdx > 7 Then colorIdx = 0
        End If

        gColorIdx(m) = colorIdx

        gDoc.Selection.Clear
        gDoc.Selection.Add gShapes(m)
        Set oVP = gDoc.Selection.VisProperties
        oVP.SetRealColor gColors(colorIdx * 3), _
                         gColors(colorIdx * 3 + 1), _
                         gColors(colorIdx * 3 + 2), 1
        gDoc.Selection.Clear

    Next m

    '------------------------------------------------------------
    ' STEP 8 - Select all found fillets in CATIA
    '------------------------------------------------------------
    gDoc.Selection.Clear
    For m = 0 To gCount - 1
        gDoc.Selection.Add gShapes(m)
    Next m

    '------------------------------------------------------------
    ' STEP 9 - Refresh and show UserForm
    '------------------------------------------------------------
    gPart.Update
    FrmFilletReport.Show vbModeless

End Sub


'================================================================
' ReadRadius : returns radius in mm, or -1 if not readable
'================================================================
Function ReadRadius(oShape As AnyObject, oPart As Part) As Double

    Dim dRadius As Double
    Dim oParam  As Parameter
    Dim oFillet As EdgeFillet
    Dim dVal    As Double
    Dim j       As Integer

    dRadius = -1

    For j = 1 To oPart.Parameters.Count
        Set oParam = oPart.Parameters.Item(j)
        If InStr(oParam.Name, oShape.Name) > 0 Then
            If InStr(LCase(oParam.Name), "radius") > 0 Then
                dRadius = oParam.Value
                ReadRadius = dRadius
                Exit Function
            End If
        End If
    Next j

    On Error Resume Next
    Set oFillet = oShape
    If Not oFillet Is Nothing Then
        dVal = oFillet.Radius.Value
        If Err.Number = 0 Then
            dRadius = dVal
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ReadRadius = dRadius

End Function


'################################################################
' USERFORM CODE — paste into FrmFilletReport code window
' (double-click the UserForm to open its code window)
'################################################################

'Private Sub UserForm_Initialize()
'
'    Dim m          As Long
'    Dim curRad     As Double
'    Dim sLine      As String
'    Dim nGroups    As Integer
'
'    '-- Form title -------------------------------------------
'    Me.Caption = "Non-Standard Fillets — " & gDoc.Name
'
'    '-- Header label -----------------------------------------
'    Me.lblTitle.Caption = "Part: " & gDoc.Name & _
'                          "   |   Non-standard fillets found: " & gCount
'
'    '-- Populate ListBox -------------------------------------
'    '   Format of each line:
'    '   [Colour]  EdgeFillet.XX  ->  X.XXX mm
'    '   Group headers are added as separator lines
'    Me.lstFillets.Clear
'    curRad  = -1
'    nGroups = 0
'
'    For m = 0 To gCount - 1
'
'        '-- Group header line when radius changes
'        If gRadii(m) <> curRad Then
'            If curRad > 0 Then
'                Me.lstFillets.AddItem "   --- Subtotal: " & _
'                    CountGroup(curRad) & " fillet(s) ---"
'                Me.lstFillets.AddItem ""
'            End If
'            curRad  = gRadii(m)
'            nGroups = nGroups + 1
'            Me.lstFillets.AddItem _
'                "== Radius: " & Format(curRad, "0.000") & " mm" & _
'                "   Colour: " & gColorNames(gColorIdx(m)) & " =="
'        End If
'
'        '-- Fillet line
'        sLine = "   " & gNames(m) & _
'                "   ->   " & Format(gRadii(m), "0.000") & " mm" & _
'                "   [" & gColorNames(gColorIdx(m)) & "]"
'        Me.lstFillets.AddItem sLine
'
'    Next m
'
'    '-- Last group subtotal
'    Me.lstFillets.AddItem "   --- Subtotal: " & _
'        CountGroup(curRad) & " fillet(s) ---"
'
'    '-- Grand total label
'    Me.lblGrandTotal.Caption = "TOTAL non-standard fillets: " & gCount & _
'                               "   |   Radius groups: " & nGroups
'
'End Sub
'
'
'----------------------------------------------------------------
' Click on ListBox -> select that fillet in CATIA
' Header/separator lines are ignored (they don't start with spaces)
'----------------------------------------------------------------
'Private Sub lstFillets_Click()
'
'    Dim sSelected As String
'    Dim m         As Long
'
'    If Me.lstFillets.ListIndex < 0 Then Exit Sub
'
'    sSelected = Trim(Me.lstFillets.List(Me.lstFillets.ListIndex))
'
'    '-- Ignore group headers and separator lines
'    If Left(sSelected, 2) = "==" Then Exit Sub
'    If Left(sSelected, 3) = "---" Then Exit Sub
'    If Len(sSelected) = 0 Then Exit Sub
'
'    '-- Extract fillet name (text before first "->")
'    Dim arParts() As String
'    arParts = Split(sSelected, "->")
'    Dim sName As String
'    sName = Trim(arParts(0))
'
'    '-- Find matching fillet and select it in CATIA
'    For m = 0 To gCount - 1
'        If gNames(m) = sName Then
'            gDoc.Selection.Clear
'            gDoc.Selection.Add gShapes(m)
'            Exit For
'        End If
'    Next m
'
'End Sub
'
'
'----------------------------------------------------------------
' Copy button -> build report text and copy to clipboard
'----------------------------------------------------------------
'Private Sub btnCopy_Click()
'
'    Dim sReport As String
'    sReport = BuildReport()
'
'    '-- Use a temporary hidden TextBox to access clipboard
'    Dim oBox As MSForms.TextBox
'    Set oBox = Me.Controls.Add("Forms.TextBox.1")
'    oBox.Visible  = False
'    oBox.MultiLine = True
'    oBox.Text      = sReport
'    oBox.SetFocus
'    oBox.SelStart  = 0
'    oBox.SelLength = Len(oBox.Text)
'    oBox.Copy
'    Me.Controls.Remove oBox.Name
'
'    MsgBox "Report copied to clipboard!", vbInformation, "Copied"
'
'End Sub
'
'
'----------------------------------------------------------------
' Export button -> save report as txt on Desktop
'----------------------------------------------------------------
'Private Sub btnExport_Click()
'
'    Dim sReport   As String
'    Dim sDesktop  As String
'    Dim sFilePath As String
'    Dim iFile     As Integer
'
'    sReport   = BuildReport()
'    sDesktop  = Environ("USERPROFILE") & "\Desktop\"
'    sFilePath = sDesktop & "FilletReport.txt"
'    iFile     = FreeFile
'
'    Open sFilePath For Output As #iFile
'    Print #iFile, sReport
'    Close #iFile
'
'    MsgBox "Report saved to:" & vbCrLf & sFilePath, _
'           vbInformation, "Exported"
'
'End Sub
'
'
'----------------------------------------------------------------
' Close button
'----------------------------------------------------------------
'Private Sub btnClose_Click()
'    Unload Me
'End Sub
'
'
'----------------------------------------------------------------
' BuildReport : builds the full report string
'----------------------------------------------------------------
'Private Function BuildReport() As String
'
'    Dim m          As Long
'    Dim curRad     As Double
'    Dim groupCount As Long
'    Dim sReport    As String
'
'    curRad     = -1
'    groupCount = 0
'    sReport    = "NON-STANDARD FILLETS REPORT" & vbCrLf
'    sReport    = sReport & "Part: " & gDoc.Name & vbCrLf
'    sReport    = sReport & "=============================" & vbCrLf & vbCrLf
'
'    For m = 0 To gCount - 1
'
'        If gRadii(m) <> curRad Then
'            If curRad > 0 Then
'                sReport = sReport & "  Subtotal: " & groupCount & _
'                           " fillet(s)" & vbCrLf & vbCrLf
'            End If
'            curRad     = gRadii(m)
'            groupCount = 0
'            sReport    = sReport & "-- Radius = " & Format(curRad, "0.000") & _
'                         " mm   [Colour: " & gColorNames(gColorIdx(m)) & "] --" & vbCrLf
'        End If
'
'        groupCount = groupCount + 1
'        sReport = sReport & "   " & groupCount & ". " & gNames(m) & _
'                  "   ->   " & Format(gRadii(m), "0.000") & " mm" & _
'                  "   [" & gColorNames(gColorIdx(m)) & "]" & vbCrLf
'
'    Next m
'
'    sReport = sReport & "  Subtotal: " & groupCount & " fillet(s)" & vbCrLf & vbCrLf
'    sReport = sReport & "=============================" & vbCrLf
'    sReport = sReport & "TOTAL non-standard: " & gCount & vbCrLf
'
'    BuildReport = sReport
'
'End Function
'
'
'----------------------------------------------------------------
' CountGroup : counts fillets with a given radius
'----------------------------------------------------------------
'Private Function CountGroup(dRad As Double) As Long
'
'    Dim m As Long
'    Dim n As Long
'    n = 0
'    For m = 0 To gCount - 1
'        If gRadii(m) = dRad Then n = n + 1
'    Next m
'    CountGroup = n
'
'End Function

