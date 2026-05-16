VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmLegend 
   Caption         =   "UserForm1"
   ClientHeight    =   6750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11100
   OleObjectBlob   =   "FrmLegend.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmLegend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblTitle_Click()

End Sub

'################################################################
' USERFORM CODE
' --> Create a UserForm named FrmLegend
'
' Controls needed:
'   lblTitle        Label           Bold=True, at top, wide
'   fraLegend       Frame           Width=380, Height=240, ScrollBars=2 (vertical)
'   lblSummary      Label           below frame, shows total count
'   btnCopy         CommandButton   Caption="Copy to clipboard"
'   btnExport       CommandButton   Caption="Export to .txt"
'   btnClose        CommandButton   Caption="Close"
'
' Paste everything below (removing leading ') into the
' FrmLegend code window (double-click the form to open it)
'################################################################

Private Sub UserForm_Initialize()

    Dim grp       As Long
    Dim oLbl      As MSForms.Label
    Dim oBox      As MSForms.Label
    Dim yPos      As Long
    Dim lR        As Long
    Dim lG        As Long
    Dim lB        As Long

    Me.Caption = "Fillet Colour Legend — " & gDoc.Name
    Me.lblTitle.Caption = "Fillet groups by radius — Total fillets: " & gCount & _
                          "   |   Groups: " & gGroupCount

    '-- Build one row per radius group inside fraLegend ---------
    yPos = 8

    For grp = 0 To gGroupCount - 1

        lR = gColorsR(gGroupColorIdx(grp))
        lG = gColorsG(gGroupColorIdx(grp))
        lB = gColorsB(gGroupColorIdx(grp))

        '-- Coloured rectangle (Label with background colour) --
        Set oBox = fraLegend.Controls.Add("Forms.Label.1")
        With oBox
            .Left = 6
            .Top = yPos
            .Width = 30
            .Height = 18
            .BackColor = RGB(lR, lG, lB)
            .BorderStyle = 1
            .Caption = ""
        End With

        '-- Text label with radius, colour name and count ------
        Set oLbl = fraLegend.Controls.Add("Forms.Label.1")
        With oLbl
            .Left = 44
            .Top = yPos + 1
            .Width = 320
            .Height = 16
            .Caption = Format(gGroupRadii(grp), "0.000") & " mm" & _
                       "   [" & gColorNames(gGroupColorIdx(grp)) & "]" & _
                       "   —   " & gGroupCounts(grp) & " fillet(s)"
            .Font.Size = 9
        End With

        yPos = yPos + 24

    Next grp

    '-- Resize frame height to fit all rows --------------------
    If yPos + 10 > fraLegend.Height Then
        fraLegend.ScrollHeight = yPos + 10
    End If

    Me.lblSummary.Caption = "Total: " & gCount & " fillet(s) in " & _
                            gGroupCount & " radius group(s)"

End Sub


Private Sub btnCopy_Click()
    Dim sReport As String
    sReport = BuildReport()
    '-- Place TextBox off-screen (visible but not seen) ---------
    Dim oTxt As MSForms.TextBox
    Set oTxt = Me.Controls.Add("Forms.TextBox.1")
    oTxt.Visible = True
    oTxt.MultiLine = True
    oTxt.Left = -500
    oTxt.Top = -500
    oTxt.Width = 100
    oTxt.Height = 20
    oTxt.Text = sReport
    oTxt.SetFocus
    oTxt.SelStart = 0
    oTxt.SelLength = Len(oTxt.Text)
    oTxt.Copy
    Me.Controls.Remove oTxt.Name
    MsgBox "Report copied to clipboard!", vbInformation, "Copied"
End Sub


Private Sub btnExport_Click()
    Dim sReport   As String
    Dim sDesktop  As String
    Dim sFilePath As String
    Dim iFile     As Integer
    sReport = BuildReport()
    sDesktop = Environ("USERPROFILE") & "\Desktop\"
    sFilePath = sDesktop & "FilletColourReport.txt"
    iFile = FreeFile
    Open sFilePath For Output As #iFile
    Print #iFile, sReport
    Close #iFile
    MsgBox "Saved to: " & sFilePath, vbInformation, "Exported"
End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub

