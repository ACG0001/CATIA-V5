VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmFilletReport 
   Caption         =   "UserForm1"
   ClientHeight    =   7950
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11925
   OleObjectBlob   =   "FrmFilletReport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmFilletReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim m       As Long
    Dim curRad  As Double
    Dim sLine   As String
    Dim nGroups As Integer
    Me.Caption = "Non-Standard Fillets — " & gDoc.Name
    Me.lblTitle.Caption = "Part: " & gDoc.Name & _
                          "   |   Non-standard fillets found: " & gCount
    Me.lstFillets.Clear
    curRad = -1
    nGroups = 0
    For m = 0 To gCount - 1
        If gRadii(m) <> curRad Then
            If curRad > 0 Then
                Me.lstFillets.AddItem "   --- Subtotal: " & _
                    CountGroup(curRad) & " fillet(s) ---"
                Me.lstFillets.AddItem ""
            End If
            curRad = gRadii(m)
            nGroups = nGroups + 1
            Me.lstFillets.AddItem _
                "== Radius: " & Format(curRad, "0.000") & " mm" & _
                "   Colour: " & gColorNames(gColorIdx(m)) & " =="
        End If
        sLine = "   " & gNames(m) & _
                "   ->   " & Format(gRadii(m), "0.000") & " mm" & _
                "   [" & gColorNames(gColorIdx(m)) & "]"
        Me.lstFillets.AddItem sLine
    Next m
    Me.lstFillets.AddItem "   --- Subtotal: " & _
        CountGroup(curRad) & " fillet(s) ---"
    Me.lblGrandTotal.Caption = "TOTAL non-standard fillets: " & gCount & _
                               "   |   Radius groups: " & nGroups
End Sub
Private Sub lstFillets_Click()
    Dim sSelected As String
    Dim arParts() As String
    Dim sName     As String
    Dim m         As Long
    If Me.lstFillets.ListIndex < 0 Then Exit Sub
    sSelected = Trim(Me.lstFillets.List(Me.lstFillets.ListIndex))
    If Left(sSelected, 2) = "==" Then Exit Sub
    If Left(sSelected, 3) = "---" Then Exit Sub
    If Len(sSelected) = 0 Then Exit Sub
    arParts = Split(sSelected, "->")
    sName = Trim(arParts(0))
    For m = 0 To gCount - 1
        If gNames(m) = sName Then
            gDoc.Selection.Clear
            gDoc.Selection.Add gShapes(m)
            Exit For
        End If
    Next m
End Sub
Private Sub btnCopy_Click()
    Dim sReport As String
    Dim oData   As DataObject
    sReport = BuildReport()
    Set oData = New DataObject
    oData.SetText sReport
    oData.PutInClipboard
    MsgBox "Report copied to clipboard!", vbInformation, "Copied"
End Sub
Private Sub btnExport_Click()
    Dim sReport   As String
    Dim sDesktop  As String
    Dim sFilePath As String
    Dim iFile     As Integer
    sReport = BuildReport()
    sDesktop = Environ("USERPROFILE") & "\Desktop\"
    sFilePath = sDesktop & "FilletReport.txt"
    iFile = FreeFile
    Open sFilePath For Output As #iFile
    Print #iFile, sReport
    Close #iFile
    MsgBox "Report saved to:" & vbCrLf & sFilePath, vbInformation, "Exported"
End Sub
Private Sub btnClose_Click()
    Unload Me
End Sub
Private Function BuildReport() As String
    Dim m          As Long
    Dim curRad     As Double
    Dim groupCount As Long
    Dim sReport    As String
    curRad = -1
    groupCount = 0
    sReport = "NON-STANDARD FILLETS REPORT" & vbCrLf
    sReport = sReport & "Part: " & gDoc.Name & vbCrLf
    sReport = sReport & "=============================" & vbCrLf & vbCrLf
    For m = 0 To gCount - 1
        If gRadii(m) <> curRad Then
            If curRad > 0 Then
                sReport = sReport & "  Subtotal: " & groupCount & _
                           " fillet(s)" & vbCrLf & vbCrLf
            End If
            curRad = gRadii(m)
            groupCount = 0
            sReport = sReport & "-- Radius = " & Format(curRad, "0.000") & _
                         " mm   [Colour: " & gColorNames(gColorIdx(m)) & "] --" & vbCrLf
        End If
        groupCount = groupCount + 1
        sReport = sReport & "   " & groupCount & ". " & gNames(m) & _
                  "   ->   " & Format(gRadii(m), "0.000") & " mm" & _
                  "   [" & gColorNames(gColorIdx(m)) & "]" & vbCrLf
    Next m
    sReport = sReport & "  Subtotal: " & groupCount & " fillet(s)" & vbCrLf & vbCrLf
    sReport = sReport & "=============================" & vbCrLf
    sReport = sReport & "TOTAL non-standard: " & gCount & vbCrLf
    BuildReport = sReport
End Function
Private Function CountGroup(dRad As Double) As Long
    Dim m As Long
    Dim n As Long
    n = 0
    For m = 0 To gCount - 1
        If gRadii(m) = dRad Then n = n + 1
    Next m
    CountGroup = n
End Function

