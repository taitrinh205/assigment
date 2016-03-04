<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.HệThốngToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ThoátToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Thoát2ToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuảnLýToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.SảnPhẩmToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.KháchHàngToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HỗTrợToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.HướngDẫnSửDụngToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ThôngTinPhầnMềmToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.PictureBox2 = New System.Windows.Forms.PictureBox()
        Me.lblTenSV = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.FromArgb(CType(CType(75, Byte), Integer), CType(CType(25, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.MenuStrip1.Font = New System.Drawing.Font("Segoe UI Semibold", 11.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HệThốngToolStripMenuItem, Me.QuảnLýToolStripMenuItem, Me.HỗTrợToolStripMenuItem1})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(686, 28)
        Me.MenuStrip1.TabIndex = 0
        '
        'HệThốngToolStripMenuItem
        '
        Me.HệThốngToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ThoátToolStripMenuItem, Me.Thoát2ToolStripMenuItem1})
        Me.HệThốngToolStripMenuItem.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.HệThốngToolStripMenuItem.Name = "HệThốngToolStripMenuItem"
        Me.HệThốngToolStripMenuItem.Size = New System.Drawing.Size(96, 24)
        Me.HệThốngToolStripMenuItem.Text = "HỆ THỐNG"
        '
        'ThoátToolStripMenuItem
        '
        Me.ThoátToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(74, Byte), Integer), CType(CType(25, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.ThoátToolStripMenuItem.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.ThoátToolStripMenuItem.Name = "ThoátToolStripMenuItem"
        Me.ThoátToolStripMenuItem.Size = New System.Drawing.Size(150, 24)
        Me.ThoátToolStripMenuItem.Text = "Đăng Xuất"
        '
        'Thoát2ToolStripMenuItem1
        '
        Me.Thoát2ToolStripMenuItem1.BackColor = System.Drawing.Color.FromArgb(CType(CType(74, Byte), Integer), CType(CType(25, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.Thoát2ToolStripMenuItem1.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.Thoát2ToolStripMenuItem1.Name = "Thoát2ToolStripMenuItem1"
        Me.Thoát2ToolStripMenuItem1.Size = New System.Drawing.Size(150, 24)
        Me.Thoát2ToolStripMenuItem1.Text = "Thoát"
        '
        'QuảnLýToolStripMenuItem
        '
        Me.QuảnLýToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(74, Byte), Integer), CType(CType(25, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.QuảnLýToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SảnPhẩmToolStripMenuItem, Me.KháchHàngToolStripMenuItem})
        Me.QuảnLýToolStripMenuItem.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.QuảnLýToolStripMenuItem.Name = "QuảnLýToolStripMenuItem"
        Me.QuảnLýToolStripMenuItem.Size = New System.Drawing.Size(85, 24)
        Me.QuảnLýToolStripMenuItem.Text = "QUẢN LÝ"
        '
        'SảnPhẩmToolStripMenuItem
        '
        Me.SảnPhẩmToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(74, Byte), Integer), CType(CType(25, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.SảnPhẩmToolStripMenuItem.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.SảnPhẩmToolStripMenuItem.Name = "SảnPhẩmToolStripMenuItem"
        Me.SảnPhẩmToolStripMenuItem.Size = New System.Drawing.Size(161, 24)
        Me.SảnPhẩmToolStripMenuItem.Text = "Sản Phẩm"
        '
        'KháchHàngToolStripMenuItem
        '
        Me.KháchHàngToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(74, Byte), Integer), CType(CType(25, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.KháchHàngToolStripMenuItem.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.KháchHàngToolStripMenuItem.Name = "KháchHàngToolStripMenuItem"
        Me.KháchHàngToolStripMenuItem.Size = New System.Drawing.Size(161, 24)
        Me.KháchHàngToolStripMenuItem.Text = "Khách Hàng"
        '
        'HỗTrợToolStripMenuItem1
        '
        Me.HỗTrợToolStripMenuItem1.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HướngDẫnSửDụngToolStripMenuItem, Me.ThôngTinPhầnMềmToolStripMenuItem})
        Me.HỗTrợToolStripMenuItem1.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.HỗTrợToolStripMenuItem1.ImageTransparentColor = System.Drawing.Color.Transparent
        Me.HỗTrợToolStripMenuItem1.Name = "HỗTrợToolStripMenuItem1"
        Me.HỗTrợToolStripMenuItem1.Size = New System.Drawing.Size(76, 24)
        Me.HỗTrợToolStripMenuItem1.Text = "HỖ TRỢ"
        '
        'HướngDẫnSửDụngToolStripMenuItem
        '
        Me.HướngDẫnSửDụngToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(74, Byte), Integer), CType(CType(25, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.HướngDẫnSửDụngToolStripMenuItem.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.HướngDẫnSửDụngToolStripMenuItem.Name = "HướngDẫnSửDụngToolStripMenuItem"
        Me.HướngDẫnSửDụngToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.HướngDẫnSửDụngToolStripMenuItem.Text = "Hướng dẫn sử dụng"
        '
        'ThôngTinPhầnMềmToolStripMenuItem
        '
        Me.ThôngTinPhầnMềmToolStripMenuItem.BackColor = System.Drawing.Color.FromArgb(CType(CType(74, Byte), Integer), CType(CType(25, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.ThôngTinPhầnMềmToolStripMenuItem.ForeColor = System.Drawing.Color.DarkTurquoise
        Me.ThôngTinPhầnMềmToolStripMenuItem.Name = "ThôngTinPhầnMềmToolStripMenuItem"
        Me.ThôngTinPhầnMềmToolStripMenuItem.Size = New System.Drawing.Size(221, 24)
        Me.ThôngTinPhầnMềmToolStripMenuItem.Text = "Thông tin phần mềm"
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Image = Global.INF205_Assignment.My.Resources.Resources.poly
        Me.PictureBox1.Location = New System.Drawing.Point(12, 331)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(326, 90)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 1
        Me.PictureBox1.TabStop = False
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.Image = CType(resources.GetObject("PictureBox2.Image"), System.Drawing.Image)
        Me.PictureBox2.Location = New System.Drawing.Point(408, 331)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(261, 90)
        Me.PictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox2.TabIndex = 2
        Me.PictureBox2.TabStop = False
        '
        'lblTenSV
        '
        Me.lblTenSV.BackColor = System.Drawing.Color.Transparent
        Me.lblTenSV.Font = New System.Drawing.Font("Segoe UI Semibold", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(163, Byte))
        Me.lblTenSV.ForeColor = System.Drawing.SystemColors.Highlight
        Me.lblTenSV.Location = New System.Drawing.Point(513, 0)
        Me.lblTenSV.Name = "lblTenSV"
        Me.lblTenSV.Size = New System.Drawing.Size(173, 46)
        Me.lblTenSV.TabIndex = 3
        Me.lblTenSV.Text = "Trịnh Nguyễn Thái Tài-Ps02714"
        Me.lblTenSV.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.INF205_Assignment.My.Resources.Resources.wallpaper_approved3
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(686, 434)
        Me.Controls.Add(Me.lblTenSV)
        Me.Controls.Add(Me.PictureBox2)
        Me.Controls.Add(Me.PictureBox1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "frmMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ứng Dụng Quản Lý Bán Hàng"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents HệThốngToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ThoátToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QuảnLýToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents SảnPhẩmToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents KháchHàngToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HỗTrợToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HướngDẫnSửDụngToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ThôngTinPhầnMềmToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Thoát2ToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents lblTenSV As System.Windows.Forms.Label
End Class
