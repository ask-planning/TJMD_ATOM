Attribute VB_Name = "Form_F_4_輸入全般_進行状況確認_SUB"
Attribute VB_Base = "0{F856006F-C15B-47DC-A945-E2301E2E72C1}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub btn_select_Click()

    If Not IsNull(Me![pk_invoice_no_main]) Then
        Call shinko_jyokyo_display(Me![pk_invoice_no_main])
    End If
    
End Sub
