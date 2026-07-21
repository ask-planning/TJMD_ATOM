<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.hdrHeader = New Atom.App.Common.HeaderControl()
        Me.flowMain = New System.Windows.Forms.FlowLayoutPanel()
        Me.grpShanghai = New System.Windows.Forms.GroupBox()
        Me.flowShanghai = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnInvoiceImport = New System.Windows.Forms.Button()
        Me.btnInvoiceCheck = New System.Windows.Forms.Button()
        Me.btnInvoiceHeader = New System.Windows.Forms.Button()
        Me.btnR3PoImport = New System.Windows.Forms.Button()
        Me.btnR3PoCheck = New System.Windows.Forms.Button()
        Me.btnR3StockTransfer = New System.Windows.Forms.Button()
        Me.btnSapAmount = New System.Windows.Forms.Button()
        Me.grpOther = New System.Windows.Forms.GroupBox()
        Me.flowOther = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnProcPoImport = New System.Windows.Forms.Button()
        Me.btnProcPoCheck = New System.Windows.Forms.Button()
        Me.btnProcDelivery = New System.Windows.Forms.Button()
        Me.grpMaster = New System.Windows.Forms.GroupBox()
        Me.flowMaster = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnRefNumber = New System.Windows.Forms.Button()
        Me.btnItemMaster = New System.Windows.Forms.Button()
        Me.btnCalendar = New System.Windows.Forms.Button()
        Me.btnSupplierMaster = New System.Windows.Forms.Button()
        Me.btnProductName = New System.Windows.Forms.Button()
        Me.btnExpenseCode = New System.Windows.Forms.Button()
        Me.grpExpense = New System.Windows.Forms.GroupBox()
        Me.flowExpense = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnMiscExpense = New System.Windows.Forms.Button()
        Me.grpOthers = New System.Windows.Forms.GroupBox()
        Me.flowOthers = New System.Windows.Forms.FlowLayoutPanel()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.flowMain.SuspendLayout()
        Me.grpShanghai.SuspendLayout()
        Me.flowShanghai.SuspendLayout()
        Me.grpOther.SuspendLayout()
        Me.flowOther.SuspendLayout()
        Me.grpMaster.SuspendLayout()
        Me.flowMaster.SuspendLayout()
        Me.grpExpense.SuspendLayout()
        Me.flowExpense.SuspendLayout()
        Me.grpOthers.SuspendLayout()
        Me.flowOthers.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnInvoiceImport
        '
        Me.btnInvoiceImport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnInvoiceImport.Name = "btnInvoiceImport"
        Me.btnInvoiceImport.Size = New System.Drawing.Size(210, 34)
        Me.btnInvoiceImport.TabIndex = 0
        Me.btnInvoiceImport.Text = "INVOICE取込"
        Me.btnInvoiceImport.UseVisualStyleBackColor = True
        Me.btnInvoiceImport.Tag = "101"
        '
        'btnInvoiceCheck
        '
        Me.btnInvoiceCheck.Margin = New System.Windows.Forms.Padding(4)
        Me.btnInvoiceCheck.Name = "btnInvoiceCheck"
        Me.btnInvoiceCheck.Size = New System.Drawing.Size(210, 34)
        Me.btnInvoiceCheck.TabIndex = 1
        Me.btnInvoiceCheck.Text = "INVOICE確認"
        Me.btnInvoiceCheck.UseVisualStyleBackColor = True
        Me.btnInvoiceCheck.Tag = "102"
        '
        'btnInvoiceHeader
        '
        Me.btnInvoiceHeader.Margin = New System.Windows.Forms.Padding(4)
        Me.btnInvoiceHeader.Name = "btnInvoiceHeader"
        Me.btnInvoiceHeader.Size = New System.Drawing.Size(210, 34)
        Me.btnInvoiceHeader.TabIndex = 2
        Me.btnInvoiceHeader.Text = "INVOICEヘッダー入力"
        Me.btnInvoiceHeader.UseVisualStyleBackColor = True
        Me.btnInvoiceHeader.Tag = "103"
        '
        'btnR3PoImport
        '
        Me.btnR3PoImport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnR3PoImport.Name = "btnR3PoImport"
        Me.btnR3PoImport.Size = New System.Drawing.Size(210, 34)
        Me.btnR3PoImport.TabIndex = 3
        Me.btnR3PoImport.Text = "R3購買発注取込"
        Me.btnR3PoImport.UseVisualStyleBackColor = True
        Me.btnR3PoImport.Tag = "104"
        '
        'btnR3PoCheck
        '
        Me.btnR3PoCheck.Margin = New System.Windows.Forms.Padding(4)
        Me.btnR3PoCheck.Name = "btnR3PoCheck"
        Me.btnR3PoCheck.Size = New System.Drawing.Size(210, 34)
        Me.btnR3PoCheck.TabIndex = 4
        Me.btnR3PoCheck.Text = "R3購買発注確認"
        Me.btnR3PoCheck.UseVisualStyleBackColor = True
        Me.btnR3PoCheck.Tag = "105"
        '
        'btnR3StockTransfer
        '
        Me.btnR3StockTransfer.Margin = New System.Windows.Forms.Padding(4)
        Me.btnR3StockTransfer.Name = "btnR3StockTransfer"
        Me.btnR3StockTransfer.Size = New System.Drawing.Size(210, 34)
        Me.btnR3StockTransfer.TabIndex = 5
        Me.btnR3StockTransfer.Text = "R3在庫転送作成"
        Me.btnR3StockTransfer.UseVisualStyleBackColor = True
        Me.btnR3StockTransfer.Tag = "106"
        '
        'btnSapAmount
        '
        Me.btnSapAmount.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSapAmount.Name = "btnSapAmount"
        Me.btnSapAmount.Size = New System.Drawing.Size(210, 34)
        Me.btnSapAmount.TabIndex = 6
        Me.btnSapAmount.Text = "SAP伝票金額入力"
        Me.btnSapAmount.UseVisualStyleBackColor = True
        Me.btnSapAmount.Tag = "107"
        '
        'btnProcPoImport
        '
        Me.btnProcPoImport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnProcPoImport.Name = "btnProcPoImport"
        Me.btnProcPoImport.Size = New System.Drawing.Size(210, 34)
        Me.btnProcPoImport.TabIndex = 7
        Me.btnProcPoImport.Text = "調達R3購買発注取込"
        Me.btnProcPoImport.UseVisualStyleBackColor = True
        Me.btnProcPoImport.Tag = "201"
        '
        'btnProcPoCheck
        '
        Me.btnProcPoCheck.Margin = New System.Windows.Forms.Padding(4)
        Me.btnProcPoCheck.Name = "btnProcPoCheck"
        Me.btnProcPoCheck.Size = New System.Drawing.Size(210, 34)
        Me.btnProcPoCheck.TabIndex = 8
        Me.btnProcPoCheck.Text = "調達R3購買発注確認"
        Me.btnProcPoCheck.UseVisualStyleBackColor = True
        Me.btnProcPoCheck.Tag = "202"
        '
        'btnProcDelivery
        '
        Me.btnProcDelivery.Margin = New System.Windows.Forms.Padding(4)
        Me.btnProcDelivery.Name = "btnProcDelivery"
        Me.btnProcDelivery.Size = New System.Drawing.Size(210, 34)
        Me.btnProcDelivery.TabIndex = 9
        Me.btnProcDelivery.Text = "調達配送手配書作成"
        Me.btnProcDelivery.UseVisualStyleBackColor = True
        Me.btnProcDelivery.Tag = "203"
        '
        'btnRefNumber
        '
        Me.btnRefNumber.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRefNumber.Name = "btnRefNumber"
        Me.btnRefNumber.Size = New System.Drawing.Size(210, 34)
        Me.btnRefNumber.TabIndex = 10
        Me.btnRefNumber.Text = "社内参照番号発番"
        Me.btnRefNumber.UseVisualStyleBackColor = True
        Me.btnRefNumber.Tag = "902"
        '
        'btnItemMaster
        '
        Me.btnItemMaster.Margin = New System.Windows.Forms.Padding(4)
        Me.btnItemMaster.Name = "btnItemMaster"
        Me.btnItemMaster.Size = New System.Drawing.Size(210, 34)
        Me.btnItemMaster.TabIndex = 11
        Me.btnItemMaster.Text = "項目マスタ"
        Me.btnItemMaster.UseVisualStyleBackColor = True
        Me.btnItemMaster.Tag = "903"
        '
        'btnCalendar
        '
        Me.btnCalendar.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCalendar.Name = "btnCalendar"
        Me.btnCalendar.Size = New System.Drawing.Size(210, 34)
        Me.btnCalendar.TabIndex = 12
        Me.btnCalendar.Text = "カレンダー"
        Me.btnCalendar.UseVisualStyleBackColor = True
        Me.btnCalendar.Tag = "904"
        '
        'btnSupplierMaster
        '
        Me.btnSupplierMaster.Margin = New System.Windows.Forms.Padding(4)
        Me.btnSupplierMaster.Name = "btnSupplierMaster"
        Me.btnSupplierMaster.Size = New System.Drawing.Size(210, 34)
        Me.btnSupplierMaster.TabIndex = 13
        Me.btnSupplierMaster.Text = "仕入先マスタ登録"
        Me.btnSupplierMaster.UseVisualStyleBackColor = True
        Me.btnSupplierMaster.Tag = "905"
        '
        'btnProductName
        '
        Me.btnProductName.Margin = New System.Windows.Forms.Padding(4)
        Me.btnProductName.Name = "btnProductName"
        Me.btnProductName.Size = New System.Drawing.Size(210, 34)
        Me.btnProductName.TabIndex = 14
        Me.btnProductName.Text = "商品名称入力"
        Me.btnProductName.UseVisualStyleBackColor = True
        Me.btnProductName.Tag = "906"
        '
        'btnExpenseCode
        '
        Me.btnExpenseCode.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExpenseCode.Name = "btnExpenseCode"
        Me.btnExpenseCode.Size = New System.Drawing.Size(210, 34)
        Me.btnExpenseCode.TabIndex = 15
        Me.btnExpenseCode.Text = "経費コード登録"
        Me.btnExpenseCode.UseVisualStyleBackColor = True
        Me.btnExpenseCode.Tag = "907"
        '
        'btnMiscExpense
        '
        Me.btnMiscExpense.Margin = New System.Windows.Forms.Padding(4)
        Me.btnMiscExpense.Name = "btnMiscExpense"
        Me.btnMiscExpense.Size = New System.Drawing.Size(210, 34)
        Me.btnMiscExpense.TabIndex = 16
        Me.btnMiscExpense.Text = "諸経費入力"
        Me.btnMiscExpense.UseVisualStyleBackColor = True
        Me.btnMiscExpense.Tag = "301"
        '
        'btnExit
        '
        Me.btnExit.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(210, 40)
        Me.btnExit.TabIndex = 17
        Me.btnExit.Text = "終　了"
        Me.btnExit.UseVisualStyleBackColor = True
        '
        'flowShanghai
        '
        Me.flowShanghai.AutoSize = True
        Me.flowShanghai.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.flowShanghai.Dock = System.Windows.Forms.DockStyle.Top
        Me.flowShanghai.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.flowShanghai.Controls.Add(Me.btnInvoiceImport)
        Me.flowShanghai.Controls.Add(Me.btnInvoiceCheck)
        Me.flowShanghai.Controls.Add(Me.btnInvoiceHeader)
        Me.flowShanghai.Controls.Add(Me.btnR3PoImport)
        Me.flowShanghai.Controls.Add(Me.btnR3PoCheck)
        Me.flowShanghai.Controls.Add(Me.btnR3StockTransfer)
        Me.flowShanghai.Controls.Add(Me.btnSapAmount)
        Me.flowShanghai.Location = New System.Drawing.Point(8, 20)
        Me.flowShanghai.Name = "flowShanghai"
        Me.flowShanghai.Size = New System.Drawing.Size(220, 40)
        Me.flowShanghai.TabIndex = 0
        Me.flowShanghai.WrapContents = False
        '
        'flowOther
        '
        Me.flowOther.AutoSize = True
        Me.flowOther.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.flowOther.Dock = System.Windows.Forms.DockStyle.Top
        Me.flowOther.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.flowOther.Controls.Add(Me.btnProcPoImport)
        Me.flowOther.Controls.Add(Me.btnProcPoCheck)
        Me.flowOther.Controls.Add(Me.btnProcDelivery)
        Me.flowOther.Location = New System.Drawing.Point(8, 20)
        Me.flowOther.Name = "flowOther"
        Me.flowOther.Size = New System.Drawing.Size(220, 40)
        Me.flowOther.TabIndex = 0
        Me.flowOther.WrapContents = False
        '
        'flowMaster
        '
        Me.flowMaster.AutoSize = True
        Me.flowMaster.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.flowMaster.Dock = System.Windows.Forms.DockStyle.Top
        Me.flowMaster.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.flowMaster.Controls.Add(Me.btnRefNumber)
        Me.flowMaster.Controls.Add(Me.btnItemMaster)
        Me.flowMaster.Controls.Add(Me.btnCalendar)
        Me.flowMaster.Controls.Add(Me.btnSupplierMaster)
        Me.flowMaster.Controls.Add(Me.btnProductName)
        Me.flowMaster.Controls.Add(Me.btnExpenseCode)
        Me.flowMaster.Location = New System.Drawing.Point(8, 20)
        Me.flowMaster.Name = "flowMaster"
        Me.flowMaster.Size = New System.Drawing.Size(220, 40)
        Me.flowMaster.TabIndex = 0
        Me.flowMaster.WrapContents = False
        '
        'flowExpense
        '
        Me.flowExpense.AutoSize = True
        Me.flowExpense.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.flowExpense.Dock = System.Windows.Forms.DockStyle.Top
        Me.flowExpense.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.flowExpense.Controls.Add(Me.btnMiscExpense)
        Me.flowExpense.Location = New System.Drawing.Point(8, 20)
        Me.flowExpense.Name = "flowExpense"
        Me.flowExpense.Size = New System.Drawing.Size(220, 40)
        Me.flowExpense.TabIndex = 0
        Me.flowExpense.WrapContents = False
        '
        'flowOthers
        '
        Me.flowOthers.AutoSize = True
        Me.flowOthers.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.flowOthers.Dock = System.Windows.Forms.DockStyle.Top
        Me.flowOthers.FlowDirection = System.Windows.Forms.FlowDirection.TopDown
        Me.flowOthers.Controls.Add(Me.btnExit)
        Me.flowOthers.Location = New System.Drawing.Point(8, 20)
        Me.flowOthers.Name = "flowOthers"
        Me.flowOthers.Size = New System.Drawing.Size(220, 40)
        Me.flowOthers.TabIndex = 0
        Me.flowOthers.WrapContents = False
        '
        'grpShanghai
        '
        Me.grpShanghai.AutoSize = True
        Me.grpShanghai.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.grpShanghai.Controls.Add(Me.flowShanghai)
        Me.grpShanghai.Margin = New System.Windows.Forms.Padding(10)
        Me.grpShanghai.MinimumSize = New System.Drawing.Size(236, 0)
        Me.grpShanghai.Name = "grpShanghai"
        Me.grpShanghai.Padding = New System.Windows.Forms.Padding(8)
        Me.grpShanghai.Size = New System.Drawing.Size(236, 64)
        Me.grpShanghai.TabIndex = 0
        Me.grpShanghai.TabStop = False
        Me.grpShanghai.Text = "上海輸入定期便＆AIR便"
        '
        'grpOther
        '
        Me.grpOther.AutoSize = True
        Me.grpOther.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.grpOther.Controls.Add(Me.flowOther)
        Me.grpOther.Margin = New System.Windows.Forms.Padding(10)
        Me.grpOther.MinimumSize = New System.Drawing.Size(236, 0)
        Me.grpOther.Name = "grpOther"
        Me.grpOther.Padding = New System.Windows.Forms.Padding(8)
        Me.grpOther.Size = New System.Drawing.Size(236, 64)
        Me.grpOther.TabIndex = 0
        Me.grpOther.TabStop = False
        Me.grpOther.Text = "その他輸入関連"
        '
        'grpMaster
        '
        Me.grpMaster.AutoSize = True
        Me.grpMaster.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.grpMaster.Controls.Add(Me.flowMaster)
        Me.grpMaster.Margin = New System.Windows.Forms.Padding(10)
        Me.grpMaster.MinimumSize = New System.Drawing.Size(236, 0)
        Me.grpMaster.Name = "grpMaster"
        Me.grpMaster.Padding = New System.Windows.Forms.Padding(8)
        Me.grpMaster.Size = New System.Drawing.Size(236, 64)
        Me.grpMaster.TabIndex = 0
        Me.grpMaster.TabStop = False
        Me.grpMaster.Text = "マスタ関連"
        '
        'grpExpense
        '
        Me.grpExpense.AutoSize = True
        Me.grpExpense.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.grpExpense.Controls.Add(Me.flowExpense)
        Me.grpExpense.Margin = New System.Windows.Forms.Padding(10)
        Me.grpExpense.MinimumSize = New System.Drawing.Size(236, 0)
        Me.grpExpense.Name = "grpExpense"
        Me.grpExpense.Padding = New System.Windows.Forms.Padding(8)
        Me.grpExpense.Size = New System.Drawing.Size(236, 64)
        Me.grpExpense.TabIndex = 0
        Me.grpExpense.TabStop = False
        Me.grpExpense.Text = "輸入経費入力"
        '
        'grpOthers
        '
        Me.grpOthers.AutoSize = True
        Me.grpOthers.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.grpOthers.Controls.Add(Me.flowOthers)
        Me.grpOthers.Margin = New System.Windows.Forms.Padding(10)
        Me.grpOthers.MinimumSize = New System.Drawing.Size(236, 0)
        Me.grpOthers.Name = "grpOthers"
        Me.grpOthers.Padding = New System.Windows.Forms.Padding(8)
        Me.grpOthers.Size = New System.Drawing.Size(236, 64)
        Me.grpOthers.TabIndex = 0
        Me.grpOthers.TabStop = False
        Me.grpOthers.Text = "OTHERS"
        '
        'flowMain
        '
        Me.flowMain.AutoScroll = True
        Me.flowMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.flowMain.Controls.Add(Me.grpShanghai)
        Me.flowMain.Controls.Add(Me.grpOther)
        Me.flowMain.Controls.Add(Me.grpMaster)
        Me.flowMain.Controls.Add(Me.grpExpense)
        Me.flowMain.Controls.Add(Me.grpOthers)
        Me.flowMain.Location = New System.Drawing.Point(0, 34)
        Me.flowMain.Name = "flowMain"
        Me.flowMain.Padding = New System.Windows.Forms.Padding(14)
        Me.flowMain.Size = New System.Drawing.Size(984, 577)
        Me.flowMain.TabIndex = 1
        '
        'hdrHeader
        '
        Me.hdrHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.hdrHeader.Name = "hdrHeader"
        Me.hdrHeader.Size = New System.Drawing.Size(984, 34)
        Me.hdrHeader.TabIndex = 0
        '
        'MainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(984, 611)
        Me.Controls.Add(Me.flowMain)
        Me.Controls.Add(Me.hdrHeader)
        Me.Font = New System.Drawing.Font("Meiryo UI", 9.0!)
        Me.Name = "MainForm"
        Me.Text = "ATOM メインメニュー"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.flowShanghai.ResumeLayout(False)
        Me.grpShanghai.ResumeLayout(False)
        Me.grpShanghai.PerformLayout()
        Me.flowOther.ResumeLayout(False)
        Me.grpOther.ResumeLayout(False)
        Me.grpOther.PerformLayout()
        Me.flowMaster.ResumeLayout(False)
        Me.grpMaster.ResumeLayout(False)
        Me.grpMaster.PerformLayout()
        Me.flowExpense.ResumeLayout(False)
        Me.grpExpense.ResumeLayout(False)
        Me.grpExpense.PerformLayout()
        Me.flowOthers.ResumeLayout(False)
        Me.grpOthers.ResumeLayout(False)
        Me.grpOthers.PerformLayout()
        Me.flowMain.ResumeLayout(False)
        Me.flowMain.PerformLayout()
        Me.ResumeLayout(False)
    End Sub

    Friend WithEvents hdrHeader As Atom.App.Common.HeaderControl
    Friend WithEvents flowMain As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents grpShanghai As System.Windows.Forms.GroupBox
    Friend WithEvents flowShanghai As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents btnInvoiceImport As System.Windows.Forms.Button
    Friend WithEvents btnInvoiceCheck As System.Windows.Forms.Button
    Friend WithEvents btnInvoiceHeader As System.Windows.Forms.Button
    Friend WithEvents btnR3PoImport As System.Windows.Forms.Button
    Friend WithEvents btnR3PoCheck As System.Windows.Forms.Button
    Friend WithEvents btnR3StockTransfer As System.Windows.Forms.Button
    Friend WithEvents btnSapAmount As System.Windows.Forms.Button
    Friend WithEvents grpOther As System.Windows.Forms.GroupBox
    Friend WithEvents flowOther As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents btnProcPoImport As System.Windows.Forms.Button
    Friend WithEvents btnProcPoCheck As System.Windows.Forms.Button
    Friend WithEvents btnProcDelivery As System.Windows.Forms.Button
    Friend WithEvents grpMaster As System.Windows.Forms.GroupBox
    Friend WithEvents flowMaster As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents btnRefNumber As System.Windows.Forms.Button
    Friend WithEvents btnItemMaster As System.Windows.Forms.Button
    Friend WithEvents btnCalendar As System.Windows.Forms.Button
    Friend WithEvents btnSupplierMaster As System.Windows.Forms.Button
    Friend WithEvents btnProductName As System.Windows.Forms.Button
    Friend WithEvents btnExpenseCode As System.Windows.Forms.Button
    Friend WithEvents grpExpense As System.Windows.Forms.GroupBox
    Friend WithEvents flowExpense As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents btnMiscExpense As System.Windows.Forms.Button
    Friend WithEvents grpOthers As System.Windows.Forms.GroupBox
    Friend WithEvents flowOthers As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents btnExit As System.Windows.Forms.Button
End Class
