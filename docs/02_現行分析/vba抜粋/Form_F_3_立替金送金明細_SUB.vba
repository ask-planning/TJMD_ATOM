Attribute VB_Name = "Form_F_3_立替金送金明細_SUB"
Attribute VB_Base = "0{A12E1ED4-08F7-4986-A98D-840BAFEC2FCF}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Compare Database
Option Explicit

Private Sub amount_AfterUpdate()

    If Not IsNull(Me![AMOUNT]) Then
        If IsNull(Me![syori_ymd]) Then
            MsgBox ("処理日が未入力です。")
        End If
    End If

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_3_立替金送金明細_MAIN]![login_code]
    
End Sub

Private Sub himoku_AfterUpdate()

    If Not IsNull(Me![himoku]) Then
        If IsNull(Me![syori_ymd]) Then
            MsgBox ("処理日が未入力です。")
        End If
    End If

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_3_立替金送金明細_MAIN]![login_code]
    
End Sub

Private Sub pk_invoice_no_main_AfterUpdate()

    If Not IsNull(Me![pk_invoice_no_main]) Then
        If IsNull(Me![syori_ymd]) Then
            MsgBox ("処理日が未入力です。")
        End If
    End If

    Me![update_ymd] = Now()
    Me![update_login] = Forms![F_3_立替金送金明細_MAIN]![login_code]
    
End Sub
