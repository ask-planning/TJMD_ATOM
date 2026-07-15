Attribute VB_Name = "Form_F_1_上海輸入_INVOICE_SUB"
Attribute VB_Base = "0{81FEB6C2-65CA-4CCA-922E-1880F5DCCB03}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub shanghai_code_AfterUpdate()

    If IsNull(Me![shanghai_code]) Then
        Me![R3_hinmoku_code] = Null
        Me![shanghai_code_text] = Null
        Exit Sub
    End If

    Dim strHinmokuCode As String
    Dim strShanghaiCode As String
    
    If IsNumeric(Me![shanghai_code]) Then
        strHinmokuCode = Right("000000000000000000" & Me![shanghai_code], 18)
        strShanghaiCode = Right("00000" & Me![shanghai_code], 5)
    Else
        strHinmokuCode = Me![shanghai_code]
        strShanghaiCode = Me![shanghai_code]
    End If
    Me![shanghai_code] = strShanghaiCode

    Set cmd.ActiveConnection = conn
    
    With cmd
        .CommandText = "usp_shanghai_code_henkan_get"
        .CommandType = adCmdStoredProc
        .Parameters.Refresh
        .Parameters(1) = strShanghaiCode
        .Parameters(2) = strHinmokuCode
        .Execute
    End With
    If Not IsNull(cmd.Parameters(3)) Then
        Me![R3_hinmoku_code] = cmd.Parameters(3)
        Me![shanghai_code_text] = cmd.Parameters(4)
    Else
        Me![R3_hinmoku_code] = Null
        Me![shanghai_code_text] = "R3登録無し"
    End If
    
    Set cmd = Nothing
        

End Sub

Private Sub suryo_AfterUpdate()

    If IsNull(Me![suryo]) Or IsNull(Me![unit_price]) Then
        Exit Sub
    End If
    
    Me![AMOUNT] = Me![suryo] * Me![unit_price]
    

End Sub

Private Sub unit_price_AfterUpdate()

    If IsNull(Me![suryo]) Or IsNull(Me![unit_price]) Then
        Exit Sub
    End If
    
    Me![AMOUNT] = Me![suryo] * Me![unit_price]
    
End Sub
