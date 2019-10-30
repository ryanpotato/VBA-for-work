VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SelectionParams 
   Caption         =   "Select Your Population and Sample Size"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6330
   OleObjectBlob   =   "SelectionParams.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SelectionParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bHeader_Click()

End Sub

Private Sub cancel_Click()
    Hide
    m_Cancelled = True
    Unload Me
    End
End Sub

Private Sub RefEdit1_BeforeDragOver(cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub


Private Sub Rng_BeforeDragOver(cancel As Boolean, ByVal Data As MSForms.DataObject, ByVal x As stdole.OLE_XPOS_CONTAINER, ByVal y As stdole.OLE_YPOS_CONTAINER, ByVal DragState As MSForms.fmDragState, Effect As MSForms.fmDropEffect, ByVal Shift As Integer)

End Sub




Private Sub TextBox1_Change()

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub userform_Initialize()

'PURPOSE: Position userform to center of Excel Window (important for dual monitor compatibility)
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

'Start Userform Centered inside Excel Screen (for dual monitors)
  Me.StartUpPosition = 0
  Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
  Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
  CurrentRegionGetter.SetFocus
  bHeader.Value = True

End Sub



Private Sub run_Click()
Hide
End Sub


Private Sub UserForm_QueryClose(cancel As Integer, CloseMode As Integer)
End
End Sub
