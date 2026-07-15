Attribute VB_Name = "Form_F_1_R3購買発注_SUB3"
Attribute VB_Base = "0{57703F92-21F7-4FC7-AD48-37BC238A6653}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub PO_NO_AfterUpdate()
    
    Set cmd.ActiveConnection = conn
    
    strSql = "SELECT T_上海輸入_INVOICE_明細.PO, T_上海輸入_INVOICE_HEADER.BL_DATE, T_上海輸入_INVOICE_明細.INVOICE_NO, "
    strSql = strSql & "Sum(T_上海輸入_INVOICE_明細.金額) AS INVOICE_AMOUNT FROM T_上海輸入_INVOICE_HEADER "
    strSql = strSql & "INNER JOIN T_上海輸入_INVOICE_明細 ON T_上海輸入_INVOICE_HEADER.INVOICE_NO_MAIN = T_上海輸入_INVOICE_明細.INVOICE_NO_MAIN "
    strSql = strSql & "GROUP BY T_上海輸入_INVOICE_明細.PO, T_上海輸入_INVOICE_明細.INVOICE_NO, T_上海輸入_INVOICE_HEADER.BL_DATE  "
    strSql = strSql & "HAVING T_上海輸入_INVOICE_明細.PO = '" & Me![PO_NO] & "' "

    'Set rs1 = db.OpenRecordset(strSql)
    
    If rs1.RecordCount <> 0 Then
        Me![INVOICE_AMOUNT] = rs1![INVOICE_AMOUNT].Value
        Me![BL_DATE] = rs1![BL_DATE].Value
        Me![対象年月] = Year(rs1![BL_DATE].Value) * 100 + Month(rs1![BL_DATE].Value)
        Me![INVOICE_NO] = rs1![INVOICE_NO].Value
    Else
        Me![INVOICE_AMOUNT] = 0
    End If
    
    rs1.Close
    Set cmd = Nothing


End Sub

Private Sub SAP_NO_AfterUpdate()

    If Len(Me![SAP_NO]) <> 10 Or (IsNumeric(Me![SAP_NO]) And Left(Me![SAP_NO], 2) <> "45") Then
        MsgBox ("SAP_NOが不正です。")
        Exit Sub
    End If

End Sub

Private Sub SORT_NO_Enter()

    SendKeys "^7", False
    
End Sub

