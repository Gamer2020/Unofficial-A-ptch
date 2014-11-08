VERSION 5.00
Begin VB.Form frmPatchInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Patch Info"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPatchInfo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   137
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   233
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "3000"
   Begin VB.TextBox txtPatchData 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmPatchInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyPress(KeyAscii As Integer)
Const sThis As String = "Form_KeyPress"
    
    On Error GoTo LocalHandler
    
    ' Mimic Windows' usual behaviour
    If KeyAscii = vbKeyEscape Then
        Unload Me
    End If
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select

End Sub

Private Sub Form_Load()
Const sThis As String = "Form_Load"

    On Error GoTo LocalHandler
    
    ' Localize the form
    Localize Me
    txtPatchData.Font = "Courier New"
    Exit Sub
    
LocalHandler:

    Select Case GlobalHandler(sThis, sMyName)
        Case vbRetry
            Resume
        Case vbAbort
            Quit
        Case Else
            Resume Next
    End Select
        
End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    
    ' Free the memory associated with the form
    txtPatchData.Text = vbNullString
    Set frmPatchInfo = Nothing
    
End Sub
