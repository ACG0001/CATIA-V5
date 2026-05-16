VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmParents 
   Caption         =   "UserForm1"
   ClientHeight    =   7620
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13560
   OleObjectBlob   =   "FrmParents.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmParents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################
' USERFORM SETUP
' Create a UserForm named: FrmParents
'
' Controls:
'   Control         Name          Key settings
'   -------------------------------------------------------
'   Label           lblTitle      Bold=True, wide at top
'   Label           lblCount      below title, shows format info
'   TextBox         txtParts      Width=500, Height=300
'                                 MultiLine=True
'                                 ScrollBars=3 (both H and V)
'                                 WordWrap=False
'                                 Font=Courier New, Size=8
'   CommandButton   btnCopy       Caption="Copy to clipboard"
'   CommandButton   btnExport     Caption="Export to .txt"
'   CommandButton   btnClose      Caption="Close"
'
' Paste USERFORM CODE below into the FrmParents code window
' (double-click the form to open it, remove leading ')
'################################################################

Private Sub UserForm_Initialize()

    Dim i       As Long
    Dim sAll    As String

    Me.Caption = "FindAllParents — " & CATIA.ActiveDocument.Name

    Me.lblTitle.Caption = "Product: " & CATIA.ActiveDocument.Name & _
                          "   |   Parts found: " & gPartCount

    Me.lblCount.Caption = "Format: Parent1, Parent2, ..., Part"

    '-- Build full text with one path per line ------------------
    sAll = ""
    For i = 0 To gPartCount - 1
        sAll = sAll & gPaths(i) & vbCrLf
    Next i

    Me.txtParts.Text = sAll

End Sub


Private Sub btnCopy_Click()

    Me.txtParts.SetFocus
    Me.txtParts.SelStart = 0
    Me.txtParts.SelLength = Len(Me.txtParts.Text)
    Me.txtParts.Copy
    MsgBox "Report copied to clipboard!", vbInformation, "Copied"

End Sub


Private Sub btnExport_Click()

    Dim sReport   As String
    Dim sDesktop  As String
    Dim sFilePath As String
    Dim iFile     As Integer

    sReport = BuildReport()
    sDesktop = Environ("USERPROFILE") & "\Desktop\"
    sFilePath = sDesktop & "FindAllParents.txt"
    iFile = FreeFile

    Open sFilePath For Output As #iFile
    Print #iFile, sReport
    Close #iFile

    MsgBox "Report saved to:" & vbCrLf & sFilePath, _
           vbInformation, "Exported"

End Sub


Private Sub btnClose_Click()
    Unload Me
End Sub


