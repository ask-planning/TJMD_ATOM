Attribute VB_Name = "Form_F_M_調達輸入_商品名称入力_SUB"
Attribute VB_Base = "0{5BA43B28-BCBB-406A-B70E-B29E601016CA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub pk_hinmoku_code_AfterUpdate()

    If IsNull(Me![pk_hinmoku_code]) Or Not IsNull(Me![hinmoku_code_text]) Then
        Exit Sub
    End If
    
    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_R3_hinmoku_text_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = Me![pk_hinmoku_code]
        .Execute
    End With
    
    Me![hinmoku_code_text] = cmd.Parameters(2)
    
    Set cmd = Nothing
    
End Sub
