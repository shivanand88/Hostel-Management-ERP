<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDueFeePayment
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.txtDueCharges = New System.Windows.Forms.TextBox()
        Me.txtTotalPaid = New System.Windows.Forms.TextBox()
        Me.dtpPaymentDate = New System.Windows.Forms.DateTimePicker()
        Me.txtBankChallanNumber = New System.Windows.Forms.TextBox()
        Me.txtFine = New System.Windows.Forms.TextBox()
        Me.label19 = New System.Windows.Forms.Label()
        Me.label17 = New System.Windows.Forms.Label()
        Me.label15 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.txtHosteler = New System.Windows.Forms.TextBox()
        Me.Update_record = New System.Windows.Forms.Button()
        Me.dataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Print = New System.Windows.Forms.Button()
        Me.panel1 = New System.Windows.Forms.Panel()
        Me.NewRecord = New System.Windows.Forms.Button()
        Me.Delete = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.txtTotalCharges = New System.Windows.Forms.TextBox()
        Me.txtBranch = New System.Windows.Forms.TextBox()
        Me.txtRoomNo = New System.Windows.Forms.TextBox()
        Me.txtHostelerName = New System.Windows.Forms.TextBox()
        Me.label25 = New System.Windows.Forms.Label()
        Me.groupBox1 = New System.Windows.Forms.GroupBox()
        Me.txtHostelerID = New System.Windows.Forms.TextBox()
        Me.label9 = New System.Windows.Forms.Label()
        Me.label26 = New System.Windows.Forms.Label()
        Me.label1 = New System.Windows.Forms.Label()
        Me.label21 = New System.Windows.Forms.Label()
        Me.panel3 = New System.Windows.Forms.Panel()
        Me.groupBox4 = New System.Windows.Forms.GroupBox()
        Me.button2 = New System.Windows.Forms.Button()
        Me.txtFeePaymentID = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.cmbMonth = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.label7 = New System.Windows.Forms.Label()
        Me.txtServiceCharges = New System.Windows.Forms.TextBox()
        Me.groupBox2 = New System.Windows.Forms.GroupBox()
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.pictureBox1 = New System.Windows.Forms.PictureBox()
        CType(Me.dataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.panel1.SuspendLayout()
        Me.groupBox1.SuspendLayout()
        Me.panel3.SuspendLayout()
        Me.groupBox4.SuspendLayout()
        Me.groupBox2.SuspendLayout()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Enabled = False
        Me.Label10.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(869, 42)
        Me.Label10.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(35, 18)
        Me.Label10.TabIndex = 70
        Me.Label10.Text = "Fine"
        Me.Label10.Visible = False
        '
        'txtDueCharges
        '
        Me.txtDueCharges.Enabled = False
        Me.txtDueCharges.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDueCharges.Location = New System.Drawing.Point(963, 79)
        Me.txtDueCharges.Name = "txtDueCharges"
        Me.txtDueCharges.ReadOnly = True
        Me.txtDueCharges.Size = New System.Drawing.Size(126, 22)
        Me.txtDueCharges.TabIndex = 35
        '
        'txtTotalPaid
        '
        Me.txtTotalPaid.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTotalPaid.Location = New System.Drawing.Point(451, 84)
        Me.txtTotalPaid.Name = "txtTotalPaid"
        Me.txtTotalPaid.Size = New System.Drawing.Size(126, 22)
        Me.txtTotalPaid.TabIndex = 2
        '
        'dtpPaymentDate
        '
        Me.dtpPaymentDate.CustomFormat = "dd/MMM/yyyy"
        Me.dtpPaymentDate.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpPaymentDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpPaymentDate.Location = New System.Drawing.Point(704, 40)
        Me.dtpPaymentDate.Name = "dtpPaymentDate"
        Me.dtpPaymentDate.Size = New System.Drawing.Size(139, 24)
        Me.dtpPaymentDate.TabIndex = 1
        '
        'txtBankChallanNumber
        '
        Me.txtBankChallanNumber.Location = New System.Drawing.Point(704, 79)
        Me.txtBankChallanNumber.Name = "txtBankChallanNumber"
        Me.txtBankChallanNumber.Size = New System.Drawing.Size(139, 24)
        Me.txtBankChallanNumber.TabIndex = 3
        '
        'txtFine
        '
        Me.txtFine.Enabled = False
        Me.txtFine.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFine.Location = New System.Drawing.Point(963, 39)
        Me.txtFine.Name = "txtFine"
        Me.txtFine.Size = New System.Drawing.Size(126, 22)
        Me.txtFine.TabIndex = 30
        Me.txtFine.Visible = False
        '
        'label19
        '
        Me.label19.AutoSize = True
        Me.label19.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label19.Location = New System.Drawing.Point(852, 81)
        Me.label19.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label19.Name = "label19"
        Me.label19.Size = New System.Drawing.Size(110, 18)
        Me.label19.TabIndex = 69
        Me.label19.Text = "Due Charges left"
        '
        'label17
        '
        Me.label17.AutoSize = True
        Me.label17.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label17.Location = New System.Drawing.Point(339, 86)
        Me.label17.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label17.Name = "label17"
        Me.label17.Size = New System.Drawing.Size(70, 18)
        Me.label17.TabIndex = 67
        Me.label17.Text = "Total Paid"
        '
        'label15
        '
        Me.label15.AutoSize = True
        Me.label15.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label15.Location = New System.Drawing.Point(604, 42)
        Me.label15.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label15.Name = "label15"
        Me.label15.Size = New System.Drawing.Size(93, 18)
        Me.label15.TabIndex = 65
        Me.label15.Text = "Payment Date"
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.Location = New System.Drawing.Point(599, 84)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(100, 17)
        Me.Label11.TabIndex = 71
        Me.Label11.Text = "Bank Challan No"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.BackColor = System.Drawing.Color.Transparent
        Me.Label5.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.ForeColor = System.Drawing.SystemColors.ButtonHighlight
        Me.Label5.Location = New System.Drawing.Point(52, 323)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(110, 18)
        Me.Label5.TabIndex = 78
        Me.Label5.Text = "Search Hostelers"
        '
        'txtHosteler
        '
        Me.txtHosteler.Location = New System.Drawing.Point(168, 323)
        Me.txtHosteler.Multiline = True
        Me.txtHosteler.Name = "txtHosteler"
        Me.txtHosteler.Size = New System.Drawing.Size(203, 24)
        Me.txtHosteler.TabIndex = 77
        '
        'Update_record
        '
        Me.Update_record.Enabled = False
        Me.Update_record.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Update_record.Location = New System.Drawing.Point(153, 117)
        Me.Update_record.Name = "Update_record"
        Me.Update_record.Size = New System.Drawing.Size(95, 33)
        Me.Update_record.TabIndex = 3
        Me.Update_record.Text = "&Update"
        Me.Update_record.UseVisualStyleBackColor = True
        Me.Update_record.Visible = False
        '
        'dataGridView1
        '
        Me.dataGridView1.AllowUserToAddRows = False
        Me.dataGridView1.AllowUserToDeleteRows = False
        Me.dataGridView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dataGridView1.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.AllCells
        Me.dataGridView1.BackgroundColor = System.Drawing.Color.White
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle1.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dataGridView1.ColumnHeadersDefaultCellStyle = DataGridViewCellStyle1
        Me.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.dataGridView1.Location = New System.Drawing.Point(55, 353)
        Me.dataGridView1.MultiSelect = False
        Me.dataGridView1.Name = "dataGridView1"
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.SystemColors.Control
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.SystemColors.WindowText
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[True]
        Me.dataGridView1.RowHeadersDefaultCellStyle = DataGridViewCellStyle3
        Me.dataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.dataGridView1.Size = New System.Drawing.Size(614, 252)
        Me.dataGridView1.TabIndex = 76
        '
        'Print
        '
        Me.Print.Enabled = False
        Me.Print.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Print.Location = New System.Drawing.Point(93, 67)
        Me.Print.Name = "Print"
        Me.Print.Size = New System.Drawing.Size(95, 33)
        Me.Print.TabIndex = 4
        Me.Print.Text = "&Print"
        Me.Print.UseVisualStyleBackColor = True
        '
        'panel1
        '
        Me.panel1.BackColor = System.Drawing.Color.DarkGray
        Me.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panel1.Controls.Add(Me.Print)
        Me.panel1.Controls.Add(Me.Update_record)
        Me.panel1.Controls.Add(Me.NewRecord)
        Me.panel1.Controls.Add(Me.Delete)
        Me.panel1.Controls.Add(Me.btnSave)
        Me.panel1.ForeColor = System.Drawing.Color.Black
        Me.panel1.Location = New System.Drawing.Point(740, 377)
        Me.panel1.Name = "panel1"
        Me.panel1.Size = New System.Drawing.Size(283, 177)
        Me.panel1.TabIndex = 2
        '
        'NewRecord
        '
        Me.NewRecord.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewRecord.Location = New System.Drawing.Point(36, 18)
        Me.NewRecord.Name = "NewRecord"
        Me.NewRecord.Size = New System.Drawing.Size(95, 33)
        Me.NewRecord.TabIndex = 0
        Me.NewRecord.Text = "&New"
        Me.NewRecord.UseVisualStyleBackColor = True
        '
        'Delete
        '
        Me.Delete.Enabled = False
        Me.Delete.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Delete.Location = New System.Drawing.Point(36, 117)
        Me.Delete.Name = "Delete"
        Me.Delete.Size = New System.Drawing.Size(95, 33)
        Me.Delete.TabIndex = 2
        Me.Delete.Text = "&Delete"
        Me.Delete.UseVisualStyleBackColor = True
        Me.Delete.Visible = False
        '
        'btnSave
        '
        Me.btnSave.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnSave.Location = New System.Drawing.Point(153, 17)
        Me.btnSave.Name = "btnSave"
        Me.btnSave.Size = New System.Drawing.Size(95, 33)
        Me.btnSave.TabIndex = 1
        Me.btnSave.Text = "&Save"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'txtTotalCharges
        '
        Me.txtTotalCharges.Enabled = False
        Me.txtTotalCharges.Location = New System.Drawing.Point(137, 77)
        Me.txtTotalCharges.Name = "txtTotalCharges"
        Me.txtTotalCharges.ReadOnly = True
        Me.txtTotalCharges.Size = New System.Drawing.Size(152, 24)
        Me.txtTotalCharges.TabIndex = 4
        '
        'txtBranch
        '
        Me.txtBranch.Enabled = False
        Me.txtBranch.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBranch.Location = New System.Drawing.Point(116, 67)
        Me.txtBranch.Name = "txtBranch"
        Me.txtBranch.ReadOnly = True
        Me.txtBranch.Size = New System.Drawing.Size(173, 24)
        Me.txtBranch.TabIndex = 2
        '
        'txtRoomNo
        '
        Me.txtRoomNo.Enabled = False
        Me.txtRoomNo.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtRoomNo.Location = New System.Drawing.Point(407, 66)
        Me.txtRoomNo.Name = "txtRoomNo"
        Me.txtRoomNo.ReadOnly = True
        Me.txtRoomNo.Size = New System.Drawing.Size(86, 24)
        Me.txtRoomNo.TabIndex = 3
        '
        'txtHostelerName
        '
        Me.txtHostelerName.Enabled = False
        Me.txtHostelerName.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHostelerName.Location = New System.Drawing.Point(407, 27)
        Me.txtHostelerName.Name = "txtHostelerName"
        Me.txtHostelerName.ReadOnly = True
        Me.txtHostelerName.Size = New System.Drawing.Size(205, 24)
        Me.txtHostelerName.TabIndex = 1
        '
        'label25
        '
        Me.label25.AutoSize = True
        Me.label25.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label25.Location = New System.Drawing.Point(24, 69)
        Me.label25.Name = "label25"
        Me.label25.Size = New System.Drawing.Size(51, 18)
        Me.label25.TabIndex = 21
        Me.label25.Text = "Branch"
        '
        'groupBox1
        '
        Me.groupBox1.BackColor = System.Drawing.Color.Transparent
        Me.groupBox1.Controls.Add(Me.txtHostelerID)
        Me.groupBox1.Controls.Add(Me.txtBranch)
        Me.groupBox1.Controls.Add(Me.txtRoomNo)
        Me.groupBox1.Controls.Add(Me.txtHostelerName)
        Me.groupBox1.Controls.Add(Me.label25)
        Me.groupBox1.Controls.Add(Me.label9)
        Me.groupBox1.Controls.Add(Me.label26)
        Me.groupBox1.Controls.Add(Me.label1)
        Me.groupBox1.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.groupBox1.ForeColor = System.Drawing.Color.White
        Me.groupBox1.Location = New System.Drawing.Point(574, 56)
        Me.groupBox1.Name = "groupBox1"
        Me.groupBox1.Size = New System.Drawing.Size(628, 114)
        Me.groupBox1.TabIndex = 70
        Me.groupBox1.TabStop = False
        Me.groupBox1.Text = "Hosteler Details"
        '
        'txtHostelerID
        '
        Me.txtHostelerID.Enabled = False
        Me.txtHostelerID.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtHostelerID.Location = New System.Drawing.Point(116, 29)
        Me.txtHostelerID.Name = "txtHostelerID"
        Me.txtHostelerID.ReadOnly = True
        Me.txtHostelerID.Size = New System.Drawing.Size(173, 24)
        Me.txtHostelerID.TabIndex = 23
        '
        'label9
        '
        Me.label9.AutoSize = True
        Me.label9.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label9.Location = New System.Drawing.Point(299, 29)
        Me.label9.Name = "label9"
        Me.label9.Size = New System.Drawing.Size(100, 18)
        Me.label9.TabIndex = 10
        Me.label9.Text = "Hosteler Name"
        '
        'label26
        '
        Me.label26.AutoSize = True
        Me.label26.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label26.Location = New System.Drawing.Point(299, 68)
        Me.label26.Name = "label26"
        Me.label26.Size = New System.Drawing.Size(71, 18)
        Me.label26.TabIndex = 22
        Me.label26.Text = "Room No."
        '
        'label1
        '
        Me.label1.AutoSize = True
        Me.label1.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label1.Location = New System.Drawing.Point(24, 33)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(79, 18)
        Me.label1.TabIndex = 2
        Me.label1.Text = "Hosteler ID"
        '
        'label21
        '
        Me.label21.AutoSize = True
        Me.label21.BackColor = System.Drawing.Color.Transparent
        Me.label21.Font = New System.Drawing.Font("Palatino Linotype", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label21.ForeColor = System.Drawing.Color.White
        Me.label21.Location = New System.Drawing.Point(460, 4)
        Me.label21.Name = "label21"
        Me.label21.Size = New System.Drawing.Size(197, 26)
        Me.label21.TabIndex = 51
        Me.label21.Text = "DUE FEE PAYMENT"
        Me.label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'panel3
        '
        Me.panel3.BackColor = System.Drawing.Color.DarkGray
        Me.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.panel3.Controls.Add(Me.label21)
        Me.panel3.ForeColor = System.Drawing.Color.Black
        Me.panel3.Location = New System.Drawing.Point(-29, 12)
        Me.panel3.Name = "panel3"
        Me.panel3.Size = New System.Drawing.Size(1232, 30)
        Me.panel3.TabIndex = 75
        '
        'groupBox4
        '
        Me.groupBox4.BackColor = System.Drawing.Color.Transparent
        Me.groupBox4.Controls.Add(Me.button2)
        Me.groupBox4.Controls.Add(Me.txtFeePaymentID)
        Me.groupBox4.Controls.Add(Me.label2)
        Me.groupBox4.Location = New System.Drawing.Point(52, 65)
        Me.groupBox4.Name = "groupBox4"
        Me.groupBox4.Size = New System.Drawing.Size(516, 105)
        Me.groupBox4.TabIndex = 69
        Me.groupBox4.TabStop = False
        '
        'button2
        '
        Me.button2.ForeColor = System.Drawing.Color.Black
        Me.button2.Location = New System.Drawing.Point(371, 39)
        Me.button2.Name = "button2"
        Me.button2.Size = New System.Drawing.Size(28, 23)
        Me.button2.TabIndex = 69
        Me.button2.Text = "<"
        Me.button2.UseVisualStyleBackColor = True
        '
        'txtFeePaymentID
        '
        Me.txtFeePaymentID.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFeePaymentID.Location = New System.Drawing.Point(182, 39)
        Me.txtFeePaymentID.Name = "txtFeePaymentID"
        Me.txtFeePaymentID.ReadOnly = True
        Me.txtFeePaymentID.Size = New System.Drawing.Size(173, 24)
        Me.txtFeePaymentID.TabIndex = 0
        '
        'label2
        '
        Me.label2.AutoSize = True
        Me.label2.BackColor = System.Drawing.Color.Transparent
        Me.label2.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label2.ForeColor = System.Drawing.Color.White
        Me.label2.Location = New System.Drawing.Point(24, 40)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(80, 18)
        Me.label2.TabIndex = 11
        Me.label2.Text = "Payment ID"
        '
        'cmbMonth
        '
        Me.cmbMonth.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbMonth.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbMonth.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmbMonth.FormattingEnabled = True
        Me.cmbMonth.Items.AddRange(New Object() {"2010-11", "2011-12", "2012-13", "2013-14", "2014-15", "2015-16", "2016-17", "2017-18", "2018-19", "2019-20", "2020-21", "2021-22", "2022-23", "2023-24", "2024-25", "2025-26"})
        Me.cmbMonth.Location = New System.Drawing.Point(451, 39)
        Me.cmbMonth.Name = "cmbMonth"
        Me.cmbMonth.Size = New System.Drawing.Size(126, 25)
        Me.cmbMonth.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(24, 82)
        Me.Label4.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(93, 18)
        Me.Label4.TabIndex = 62
        Me.Label4.Text = "Total Charges"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(307, 42)
        Me.Label3.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(137, 18)
        Me.Label3.TabIndex = 60
        Me.Label3.Text = "Payment For the Year"
        '
        'label7
        '
        Me.label7.AutoSize = True
        Me.label7.Font = New System.Drawing.Font("Palatino Linotype", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.label7.Location = New System.Drawing.Point(32, 43)
        Me.label7.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.label7.Name = "label7"
        Me.label7.Size = New System.Drawing.Size(87, 18)
        Me.label7.TabIndex = 59
        Me.label7.Text = "Due Charges"
        '
        'txtServiceCharges
        '
        Me.txtServiceCharges.Enabled = False
        Me.txtServiceCharges.Location = New System.Drawing.Point(137, 38)
        Me.txtServiceCharges.Name = "txtServiceCharges"
        Me.txtServiceCharges.ReadOnly = True
        Me.txtServiceCharges.Size = New System.Drawing.Size(152, 24)
        Me.txtServiceCharges.TabIndex = 0
        '
        'groupBox2
        '
        Me.groupBox2.BackColor = System.Drawing.Color.Transparent
        Me.groupBox2.Controls.Add(Me.txtFine)
        Me.groupBox2.Controls.Add(Me.Label10)
        Me.groupBox2.Controls.Add(Me.txtBankChallanNumber)
        Me.groupBox2.Controls.Add(Me.txtDueCharges)
        Me.groupBox2.Controls.Add(Me.Label11)
        Me.groupBox2.Controls.Add(Me.txtTotalPaid)
        Me.groupBox2.Controls.Add(Me.txtTotalCharges)
        Me.groupBox2.Controls.Add(Me.dtpPaymentDate)
        Me.groupBox2.Controls.Add(Me.Label4)
        Me.groupBox2.Controls.Add(Me.label19)
        Me.groupBox2.Controls.Add(Me.cmbMonth)
        Me.groupBox2.Controls.Add(Me.label17)
        Me.groupBox2.Controls.Add(Me.Label3)
        Me.groupBox2.Controls.Add(Me.label15)
        Me.groupBox2.Controls.Add(Me.label7)
        Me.groupBox2.Controls.Add(Me.txtServiceCharges)
        Me.groupBox2.Font = New System.Drawing.Font("Palatino Linotype", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.groupBox2.ForeColor = System.Drawing.Color.White
        Me.groupBox2.Location = New System.Drawing.Point(52, 177)
        Me.groupBox2.Name = "groupBox2"
        Me.groupBox2.Size = New System.Drawing.Size(1150, 133)
        Me.groupBox2.TabIndex = 1
        Me.groupBox2.TabStop = False
        Me.groupBox2.Text = "Charges Details"
        '
        'Timer1
        '
        '
        'pictureBox1
        '
        Me.pictureBox1.BackgroundImage = Global.HostelManagementSystem.My.Resources.Resources.top_image
        Me.pictureBox1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.pictureBox1.Location = New System.Drawing.Point(-67, -77)
        Me.pictureBox1.Name = "pictureBox1"
        Me.pictureBox1.Size = New System.Drawing.Size(1252, 43)
        Me.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pictureBox1.TabIndex = 74
        Me.pictureBox1.TabStop = False
        '
        'frmDueFeePayment
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(1215, 617)
        Me.Controls.Add(Me.pictureBox1)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.txtHosteler)
        Me.Controls.Add(Me.dataGridView1)
        Me.Controls.Add(Me.panel1)
        Me.Controls.Add(Me.groupBox1)
        Me.Controls.Add(Me.panel3)
        Me.Controls.Add(Me.groupBox4)
        Me.Controls.Add(Me.groupBox2)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "frmDueFeePayment"
        Me.Text = "frmDueFeePayment"
        CType(Me.dataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.panel1.ResumeLayout(False)
        Me.groupBox1.ResumeLayout(False)
        Me.groupBox1.PerformLayout()
        Me.panel3.ResumeLayout(False)
        Me.panel3.PerformLayout()
        Me.groupBox4.ResumeLayout(False)
        Me.groupBox4.PerformLayout()
        Me.groupBox2.ResumeLayout(False)
        Me.groupBox2.PerformLayout()
        CType(Me.pictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Private WithEvents pictureBox1 As System.Windows.Forms.PictureBox
    Private WithEvents Label10 As System.Windows.Forms.Label
    Public WithEvents txtDueCharges As System.Windows.Forms.TextBox
    Public WithEvents txtTotalPaid As System.Windows.Forms.TextBox
    Public WithEvents dtpPaymentDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents txtBankChallanNumber As System.Windows.Forms.TextBox
    Public WithEvents txtFine As System.Windows.Forms.TextBox
    Private WithEvents label19 As System.Windows.Forms.Label
    Private WithEvents label17 As System.Windows.Forms.Label
    Private WithEvents label15 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Private WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents txtHosteler As System.Windows.Forms.TextBox
    Public WithEvents Update_record As System.Windows.Forms.Button
    Private WithEvents dataGridView1 As System.Windows.Forms.DataGridView
    Public WithEvents Print As System.Windows.Forms.Button
    Private WithEvents panel1 As System.Windows.Forms.Panel
    Public WithEvents NewRecord As System.Windows.Forms.Button
    Public WithEvents Delete As System.Windows.Forms.Button
    Public WithEvents btnSave As System.Windows.Forms.Button
    Public WithEvents txtTotalCharges As System.Windows.Forms.TextBox
    Public WithEvents txtBranch As System.Windows.Forms.TextBox
    Public WithEvents txtRoomNo As System.Windows.Forms.TextBox
    Public WithEvents txtHostelerName As System.Windows.Forms.TextBox
    Private WithEvents label25 As System.Windows.Forms.Label
    Private WithEvents groupBox1 As System.Windows.Forms.GroupBox
    Private WithEvents label9 As System.Windows.Forms.Label
    Private WithEvents label26 As System.Windows.Forms.Label
    Private WithEvents label1 As System.Windows.Forms.Label
    Friend WithEvents label21 As System.Windows.Forms.Label
    Private WithEvents panel3 As System.Windows.Forms.Panel
    Private WithEvents groupBox4 As System.Windows.Forms.GroupBox
    Private WithEvents button2 As System.Windows.Forms.Button
    Public WithEvents txtFeePaymentID As System.Windows.Forms.TextBox
    Private WithEvents label2 As System.Windows.Forms.Label
    Public WithEvents cmbMonth As System.Windows.Forms.ComboBox
    Private WithEvents Label4 As System.Windows.Forms.Label
    Private WithEvents Label3 As System.Windows.Forms.Label
    Private WithEvents label7 As System.Windows.Forms.Label
    Public WithEvents txtServiceCharges As System.Windows.Forms.TextBox
    Private WithEvents groupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Public WithEvents txtHostelerID As System.Windows.Forms.TextBox
End Class
