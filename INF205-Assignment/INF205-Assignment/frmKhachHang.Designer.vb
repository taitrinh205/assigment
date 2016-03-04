<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmKhachHang
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.btnTimKH = New System.Windows.Forms.Button()
        Me.btnXoaKH = New System.Windows.Forms.Button()
        Me.btnSuaKH = New System.Windows.Forms.Button()
        Me.btnThemKH = New System.Windows.Forms.Button()
        Me.txtDiaChi = New System.Windows.Forms.TextBox()
        Me.lblDiaChi = New System.Windows.Forms.Label()
        Me.txtTenKH = New System.Windows.Forms.TextBox()
        Me.lblTenKH = New System.Windows.Forms.Label()
        Me.txtMaKH = New System.Windows.Forms.TextBox()
        Me.lblMaKH = New System.Windows.Forms.Label()
        Me.dgvKH = New System.Windows.Forms.DataGridView()
        Me.pbxKH = New System.Windows.Forms.PictureBox()
        Me.txtSDT = New System.Windows.Forms.TextBox()
        Me.lblEmail = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        CType(Me.dgvKH, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.pbxKH, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.btnTimKH)
        Me.GroupBox1.Controls.Add(Me.btnXoaKH)
        Me.GroupBox1.Controls.Add(Me.btnSuaKH)
        Me.GroupBox1.Controls.Add(Me.btnThemKH)
        Me.GroupBox1.Controls.Add(Me.txtDiaChi)
        Me.GroupBox1.Controls.Add(Me.lblDiaChi)
        Me.GroupBox1.Controls.Add(Me.txtSDT)
        Me.GroupBox1.Controls.Add(Me.lblEmail)
        Me.GroupBox1.Controls.Add(Me.txtTenKH)
        Me.GroupBox1.Controls.Add(Me.lblTenKH)
        Me.GroupBox1.Controls.Add(Me.txtMaKH)
        Me.GroupBox1.Controls.Add(Me.lblMaKH)
        Me.GroupBox1.Font = New System.Drawing.Font("Segoe UI Semibold", 9.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.Maroon
        Me.GroupBox1.Location = New System.Drawing.Point(226, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(448, 206)
        Me.GroupBox1.TabIndex = 2
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Nhập Thông Tin Khách Hàng"
        '
        'btnTimKH
        '
        Me.btnTimKH.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnTimKH.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.btnTimKH.ForeColor = System.Drawing.Color.SteelBlue
        Me.btnTimKH.Image = Global.INF205_Assignment.My.Resources.Resources.Search
        Me.btnTimKH.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnTimKH.Location = New System.Drawing.Point(356, 121)
        Me.btnTimKH.Name = "btnTimKH"
        Me.btnTimKH.Size = New System.Drawing.Size(86, 66)
        Me.btnTimKH.TabIndex = 9
        Me.btnTimKH.Text = "Tìm"
        Me.btnTimKH.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnTimKH.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnTimKH.UseVisualStyleBackColor = True
        '
        'btnXoaKH
        '
        Me.btnXoaKH.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnXoaKH.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.btnXoaKH.ForeColor = System.Drawing.Color.DarkRed
        Me.btnXoaKH.Image = Global.INF205_Assignment.My.Resources.Resources.Delete
        Me.btnXoaKH.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnXoaKH.Location = New System.Drawing.Point(262, 121)
        Me.btnXoaKH.Name = "btnXoaKH"
        Me.btnXoaKH.Size = New System.Drawing.Size(88, 66)
        Me.btnXoaKH.TabIndex = 8
        Me.btnXoaKH.Text = "Xóa"
        Me.btnXoaKH.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnXoaKH.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnXoaKH.UseVisualStyleBackColor = True
        '
        'btnSuaKH
        '
        Me.btnSuaKH.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnSuaKH.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.btnSuaKH.ForeColor = System.Drawing.Color.DarkOrange
        Me.btnSuaKH.Image = Global.INF205_Assignment.My.Resources.Resources.Edit
        Me.btnSuaKH.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnSuaKH.Location = New System.Drawing.Point(356, 32)
        Me.btnSuaKH.Name = "btnSuaKH"
        Me.btnSuaKH.Size = New System.Drawing.Size(88, 70)
        Me.btnSuaKH.TabIndex = 7
        Me.btnSuaKH.Text = "Sửa"
        Me.btnSuaKH.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnSuaKH.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnSuaKH.UseVisualStyleBackColor = True
        '
        'btnThemKH
        '
        Me.btnThemKH.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnThemKH.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.btnThemKH.Font = New System.Drawing.Font("Segoe UI", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.btnThemKH.ForeColor = System.Drawing.Color.ForestGreen
        Me.btnThemKH.Image = Global.INF205_Assignment.My.Resources.Resources.ADD1
        Me.btnThemKH.ImageAlign = System.Drawing.ContentAlignment.TopCenter
        Me.btnThemKH.Location = New System.Drawing.Point(262, 32)
        Me.btnThemKH.Name = "btnThemKH"
        Me.btnThemKH.Size = New System.Drawing.Size(88, 70)
        Me.btnThemKH.TabIndex = 6
        Me.btnThemKH.Text = "Thêm"
        Me.btnThemKH.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        Me.btnThemKH.TextImageRelation = System.Windows.Forms.TextImageRelation.ImageAboveText
        Me.btnThemKH.UseVisualStyleBackColor = True
        '
        'txtDiaChi
        '
        Me.txtDiaChi.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDiaChi.Location = New System.Drawing.Point(101, 163)
        Me.txtDiaChi.Name = "txtDiaChi"
        Me.txtDiaChi.Size = New System.Drawing.Size(155, 23)
        Me.txtDiaChi.TabIndex = 5
        '
        'lblDiaChi
        '
        Me.lblDiaChi.Location = New System.Drawing.Point(15, 164)
        Me.lblDiaChi.Name = "lblDiaChi"
        Me.lblDiaChi.Size = New System.Drawing.Size(55, 22)
        Me.lblDiaChi.TabIndex = 0
        Me.lblDiaChi.Text = "Đại chỉ:"
        '
        'txtTenKH
        '
        Me.txtTenKH.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtTenKH.Location = New System.Drawing.Point(101, 78)
        Me.txtTenKH.Name = "txtTenKH"
        Me.txtTenKH.Size = New System.Drawing.Size(155, 23)
        Me.txtTenKH.TabIndex = 2
        '
        'lblTenKH
        '
        Me.lblTenKH.Location = New System.Drawing.Point(15, 81)
        Me.lblTenKH.Name = "lblTenKH"
        Me.lblTenKH.Size = New System.Drawing.Size(86, 22)
        Me.lblTenKH.TabIndex = 0
        Me.lblTenKH.Text = "Tên KH"
        '
        'txtMaKH
        '
        Me.txtMaKH.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtMaKH.Location = New System.Drawing.Point(101, 34)
        Me.txtMaKH.Name = "txtMaKH"
        Me.txtMaKH.Size = New System.Drawing.Size(155, 23)
        Me.txtMaKH.TabIndex = 1
        '
        'lblMaKH
        '
        Me.lblMaKH.Location = New System.Drawing.Point(15, 35)
        Me.lblMaKH.Name = "lblMaKH"
        Me.lblMaKH.Size = New System.Drawing.Size(86, 22)
        Me.lblMaKH.TabIndex = 0
        Me.lblMaKH.Text = "Mã số K.H:"
        '
        'dgvKH
        '
        Me.dgvKH.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvKH.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvKH.BackgroundColor = System.Drawing.SystemColors.ActiveBorder
        Me.dgvKH.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvKH.Location = New System.Drawing.Point(12, 227)
        Me.dgvKH.Name = "dgvKH"
        Me.dgvKH.Size = New System.Drawing.Size(662, 195)
        Me.dgvKH.TabIndex = 11
        '
        'pbxKH
        '
        Me.pbxKH.BackColor = System.Drawing.Color.Transparent
        Me.pbxKH.Image = Global.INF205_Assignment.My.Resources.Resources.FollowUp
        Me.pbxKH.Location = New System.Drawing.Point(16, 12)
        Me.pbxKH.Name = "pbxKH"
        Me.pbxKH.Size = New System.Drawing.Size(204, 206)
        Me.pbxKH.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pbxKH.TabIndex = 12
        Me.pbxKH.TabStop = False
        '
        'txtSDT
        '
        Me.txtSDT.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtSDT.Location = New System.Drawing.Point(101, 118)
        Me.txtSDT.Name = "txtSDT"
        Me.txtSDT.Size = New System.Drawing.Size(155, 23)
        Me.txtSDT.TabIndex = 4
        '
        'lblEmail
        '
        Me.lblEmail.Location = New System.Drawing.Point(15, 121)
        Me.lblEmail.Name = "lblEmail"
        Me.lblEmail.Size = New System.Drawing.Size(50, 22)
        Me.lblEmail.TabIndex = 0
        Me.lblEmail.Text = "SDT"
        '
        'frmKhachHang
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.INF205_Assignment.My.Resources.Resources.snapshot1r
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(686, 434)
        Me.Controls.Add(Me.pbxKH)
        Me.Controls.Add(Me.dgvKH)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "frmKhachHang"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Quản Lý Khách Hàng"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        CType(Me.dgvKH, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.pbxKH, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents btnTimKH As System.Windows.Forms.Button
    Friend WithEvents btnXoaKH As System.Windows.Forms.Button
    Friend WithEvents btnSuaKH As System.Windows.Forms.Button
    Friend WithEvents btnThemKH As System.Windows.Forms.Button
    Friend WithEvents txtDiaChi As System.Windows.Forms.TextBox
    Friend WithEvents lblDiaChi As System.Windows.Forms.Label
    Friend WithEvents txtTenKH As System.Windows.Forms.TextBox
    Friend WithEvents lblTenKH As System.Windows.Forms.Label
    Friend WithEvents txtMaKH As System.Windows.Forms.TextBox
    Friend WithEvents lblMaKH As System.Windows.Forms.Label
    Friend WithEvents dgvKH As System.Windows.Forms.DataGridView
    Friend WithEvents pbxKH As System.Windows.Forms.PictureBox
    Friend WithEvents txtSDT As System.Windows.Forms.TextBox
    Friend WithEvents lblEmail As System.Windows.Forms.Label
End Class
