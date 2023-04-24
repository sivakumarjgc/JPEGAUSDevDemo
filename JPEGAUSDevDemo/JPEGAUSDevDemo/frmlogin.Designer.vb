<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmlogin
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmlogin))
        Me.txt_servername = New System.Windows.Forms.TextBox()
        Me.txt_database = New System.Windows.Forms.TextBox()
        Me.but_cancel = New System.Windows.Forms.Button()
        Me.but_Connect = New System.Windows.Forms.Button()
        Me.txt_xmlfile = New System.Windows.Forms.TextBox()
        Me.but_xmlbrowser = New System.Windows.Forms.Button()
        Me.frmpic = New System.Windows.Forms.PictureBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.PictureBox1 = New System.Windows.Forms.PictureBox()
        Me.lab_Servername = New System.Windows.Forms.Label()
        Me.lab_database = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lab_reviewer = New System.Windows.Forms.Label()
        Me.cb_reviewtool = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.frmpic, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txt_servername
        '
        Me.txt_servername.Location = New System.Drawing.Point(85, 101)
        Me.txt_servername.Name = "txt_servername"
        Me.txt_servername.Size = New System.Drawing.Size(360, 20)
        Me.txt_servername.TabIndex = 3
        '
        'txt_database
        '
        Me.txt_database.Location = New System.Drawing.Point(85, 127)
        Me.txt_database.Name = "txt_database"
        Me.txt_database.Size = New System.Drawing.Size(360, 20)
        Me.txt_database.TabIndex = 5
        '
        'but_cancel
        '
        Me.but_cancel.BackColor = System.Drawing.SystemColors.Control
        Me.but_cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.but_cancel.Location = New System.Drawing.Point(12, 198)
        Me.but_cancel.Name = "but_cancel"
        Me.but_cancel.Size = New System.Drawing.Size(75, 23)
        Me.but_cancel.TabIndex = 8
        Me.but_cancel.Text = "Cancel"
        Me.but_cancel.UseVisualStyleBackColor = False
        '
        'but_Connect
        '
        Me.but_Connect.BackColor = System.Drawing.SystemColors.Control
        Me.but_Connect.Location = New System.Drawing.Point(399, 198)
        Me.but_Connect.Name = "but_Connect"
        Me.but_Connect.Size = New System.Drawing.Size(75, 23)
        Me.but_Connect.TabIndex = 9
        Me.but_Connect.Text = "Connect"
        Me.but_Connect.UseVisualStyleBackColor = False
        '
        'txt_xmlfile
        '
        Me.txt_xmlfile.BackColor = System.Drawing.Color.Wheat
        Me.txt_xmlfile.Location = New System.Drawing.Point(85, 75)
        Me.txt_xmlfile.Name = "txt_xmlfile"
        Me.txt_xmlfile.Size = New System.Drawing.Size(360, 20)
        Me.txt_xmlfile.TabIndex = 12
        '
        'but_xmlbrowser
        '
        Me.but_xmlbrowser.Location = New System.Drawing.Point(451, 72)
        Me.but_xmlbrowser.Name = "but_xmlbrowser"
        Me.but_xmlbrowser.Size = New System.Drawing.Size(23, 23)
        Me.but_xmlbrowser.TabIndex = 13
        Me.but_xmlbrowser.Text = "..."
        Me.but_xmlbrowser.UseVisualStyleBackColor = True
        '
        'frmpic
        '
        Me.frmpic.BackColor = System.Drawing.SystemColors.ActiveBorder
        Me.frmpic.Image = CType(resources.GetObject("frmpic.Image"), System.Drawing.Image)
        Me.frmpic.Location = New System.Drawing.Point(36, 446)
        Me.frmpic.Name = "frmpic"
        Me.frmpic.Size = New System.Drawing.Size(370, 64)
        Me.frmpic.TabIndex = 10
        Me.frmpic.TabStop = False
        '
        'Panel1
        '
        Me.Panel1.BackgroundImage = CType(resources.GetObject("Panel1.BackgroundImage"), System.Drawing.Image)
        Me.Panel1.Controls.Add(Me.PictureBox1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(488, 64)
        Me.Panel1.TabIndex = 14
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Image = CType(resources.GetObject("PictureBox1.Image"), System.Drawing.Image)
        Me.PictureBox1.Location = New System.Drawing.Point(163, 8)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(150, 51)
        Me.PictureBox1.TabIndex = 15
        Me.PictureBox1.TabStop = False
        '
        'lab_Servername
        '
        Me.lab_Servername.AutoSize = True
        Me.lab_Servername.Location = New System.Drawing.Point(7, 104)
        Me.lab_Servername.Name = "lab_Servername"
        Me.lab_Servername.Size = New System.Drawing.Size(72, 13)
        Me.lab_Servername.TabIndex = 0
        Me.lab_Servername.Text = "Server Name:"
        '
        'lab_database
        '
        Me.lab_database.AutoSize = True
        Me.lab_database.Location = New System.Drawing.Point(23, 130)
        Me.lab_database.Name = "lab_database"
        Me.lab_database.Size = New System.Drawing.Size(56, 13)
        Me.lab_database.TabIndex = 4
        Me.lab_database.Text = "Database:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(28, 78)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(51, 13)
        Me.Label1.TabIndex = 11
        Me.Label1.Text = "XML File:"
        '
        'lab_reviewer
        '
        Me.lab_reviewer.AutoSize = True
        Me.lab_reviewer.Location = New System.Drawing.Point(9, 161)
        Me.lab_reviewer.Name = "lab_reviewer"
        Me.lab_reviewer.Size = New System.Drawing.Size(70, 13)
        Me.lab_reviewer.TabIndex = 15
        Me.lab_reviewer.Text = "Review Tool:"
        '
        'cb_reviewtool
        '
        Me.cb_reviewtool.BackColor = System.Drawing.Color.White
        Me.cb_reviewtool.FormattingEnabled = True
        Me.cb_reviewtool.Location = New System.Drawing.Point(85, 158)
        Me.cb_reviewtool.Name = "cb_reviewtool"
        Me.cb_reviewtool.Size = New System.Drawing.Size(360, 21)
        Me.cb_reviewtool.TabIndex = 16
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.SystemColors.Control
        Me.Button1.Location = New System.Drawing.Point(273, 198)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 9
        Me.Button1.Text = "Connect"
        Me.Button1.UseVisualStyleBackColor = False
        Me.Button1.Visible = False
        '
        'frmlogin
        '
        Me.AcceptButton = Me.but_Connect
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Lavender
        Me.CancelButton = Me.but_cancel
        Me.ClientSize = New System.Drawing.Size(488, 233)
        Me.Controls.Add(Me.cb_reviewtool)
        Me.Controls.Add(Me.lab_reviewer)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.but_xmlbrowser)
        Me.Controls.Add(Me.txt_xmlfile)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.frmpic)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.but_Connect)
        Me.Controls.Add(Me.but_cancel)
        Me.Controls.Add(Me.txt_database)
        Me.Controls.Add(Me.lab_database)
        Me.Controls.Add(Me.txt_servername)
        Me.Controls.Add(Me.lab_Servername)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmlogin"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "J-MRMS"
        CType(Me.frmpic, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txt_servername As System.Windows.Forms.TextBox
    Friend WithEvents txt_database As System.Windows.Forms.TextBox
    Friend WithEvents but_cancel As System.Windows.Forms.Button
    Friend WithEvents but_Connect As System.Windows.Forms.Button
    Friend WithEvents but_xmlbrowser As System.Windows.Forms.Button
    Friend WithEvents txt_xmlfile As System.Windows.Forms.TextBox
    Friend WithEvents frmpic As System.Windows.Forms.PictureBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents lab_Servername As System.Windows.Forms.Label
    Friend WithEvents lab_database As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lab_reviewer As System.Windows.Forms.Label
    Friend WithEvents cb_reviewtool As System.Windows.Forms.ComboBox
    Friend WithEvents Button1 As Button
End Class
