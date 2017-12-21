
Imports System.Windows.Forms

Public Class frmMain
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents picMacro As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    Friend WithEvents cmdMenu As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    Friend WithEvents Label2 As Microsoft.VisualBasic.Compatibility.VB6.LabelArray
    Friend WithEvents cmdMoverHechi As Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents picMacro_0 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_1 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_2 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_3 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_4 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_5 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_6 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_7 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_8 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_9 As System.Windows.Forms.PictureBox
    Friend WithEvents picMacro_10 As System.Windows.Forms.PictureBox
    Friend WithEvents Minimap As System.Windows.Forms.PictureBox
    Friend WithEvents RecChat As System.Windows.Forms.RichTextBox
    Friend WithEvents picInv As System.Windows.Forms.PictureBox
    Friend WithEvents cmdMenu_0 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdMenu_1 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdMenu_2 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdMenu_3 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdMenu_4 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdMenu_5 As System.Windows.Forms.PictureBox
    Friend WithEvents cmdMen As System.Windows.Forms.PictureBox
    Friend WithEvents nomodoseguro As System.Windows.Forms.PictureBox
    Friend WithEvents cmdInv As System.Windows.Forms.PictureBox
    Friend WithEvents Hpshp As System.Windows.Forms.Label
    Friend WithEvents AGUAsp As System.Windows.Forms.Label
    Friend WithEvents COMIDAsp As System.Windows.Forms.Label
    Friend WithEvents lblHAM As System.Windows.Forms.Label
    Friend WithEvents lblAG As System.Windows.Forms.Label
    Friend WithEvents lblMP As System.Windows.Forms.Label
    Friend WithEvents lblExp As System.Windows.Forms.Label
    Friend WithEvents GldLbl As System.Windows.Forms.Label
    Friend WithEvents lblInvInfo As System.Windows.Forms.Label
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMain))
        Me.components = New System.ComponentModel.Container()
        Me.picMacro = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
        Me.cmdMenu = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
        Me.Label2 = New Microsoft.VisualBasic.Compatibility.VB6.LabelArray(components)
        Me.cmdMoverHechi = New Microsoft.VisualBasic.Compatibility.VB6.PictureBoxArray(components)
        Me.Command1 = New System.Windows.Forms.Button()
        Me.picMacro_0 = New System.Windows.Forms.PictureBox()
        Me.picMacro_9 = New System.Windows.Forms.PictureBox()
        Me.Minimap = New System.Windows.Forms.PictureBox()
        Me.RecChat = New System.Windows.Forms.RichTextBox()
        Me.picInv = New System.Windows.Forms.PictureBox()
        Me.cmdMenu_0 = New System.Windows.Forms.PictureBox()
        Me.cmdMenu_4 = New System.Windows.Forms.PictureBox()
        Me.cmdMen = New System.Windows.Forms.PictureBox()
        Me.nomodoseguro = New System.Windows.Forms.PictureBox()
        Me.cmdInv = New System.Windows.Forms.PictureBox()
        Me.Hpshp = New System.Windows.Forms.Label()
        Me.AGUAsp = New System.Windows.Forms.Label()
        Me.COMIDAsp = New System.Windows.Forms.Label()
        Me.lblHAM = New System.Windows.Forms.Label()
        Me.lblAG = New System.Windows.Forms.Label()
        Me.lblMP = New System.Windows.Forms.Label()
        Me.lblExp = New System.Windows.Forms.Label()
        Me.GldLbl = New System.Windows.Forms.Label()
        Me.lblInvInfo = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        CType(Me.RecChat, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        CType(Me.picMacro, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdMenu, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.cmdMoverHechi, System.ComponentModel.ISupportInitialize).BeginInit()
        '
        'Command1
        '
        Me.Command1.Name = "Command1"
        Me.Command1.Visible = False
        Me.Command1.TabIndex = 35
        Me.Command1.Location = New System.Drawing.Point(566, 32)
        Me.Command1.Size = New System.Drawing.Size(33, 33)
        Me.Command1.Text = "Command1"
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'picMacro_0
        '
        Me.picMacro_0.Name = "picMacro_0"
        Me.picMacro_0.TabIndex = 27
        Me.picMacro_0.Location = New System.Drawing.Point(16, 568)
        Me.picMacro_0.Size = New System.Drawing.Size(32, 32)
        Me.picMacro_0.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.picMacro_0.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(0, Byte))
        '
        'picMacro_9
        '
        Me.picMacro_9.Name = "picMacro_9"
        Me.picMacro_9.TabIndex = 18
        Me.picMacro_9.Location = New System.Drawing.Point(371, 568)
        Me.picMacro_9.Size = New System.Drawing.Size(32, 32)
        Me.picMacro_9.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.picMacro_9.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(0, Byte))
        '
        'Minimap
        '
        Me.Minimap.Name = "Minimap"
        Me.Minimap.TabIndex = 7
        Me.Minimap.Location = New System.Drawing.Point(688, 498)
        Me.Minimap.Size = New System.Drawing.Size(98, 95)
        Me.Minimap.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.Minimap.BackColor = System.Drawing.SystemColors.MenuText
        '
        'RecChat
        '
        Me.RecChat.Name = "RecChat"
        Me.RecChat.TabStop = False
        Me.RecChat.TabIndex = 0
        Me.ToolTip1.SetToolTip(Me.RecChat, "Mensajes del servidor")
        Me.RecChat.Location = New System.Drawing.Point(14, 11)
        Me.RecChat.Size = New System.Drawing.Size(550, 101)
        Me.RecChat.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(0, Byte))
        Me.RecChat.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.RecChat.ScrollBars = System.Windows.Forms.RichTextBoxScrollBars.Vertical
        Me.RecChat.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'picInv
        '
        Me.picInv.Name = "picInv"
        Me.picInv.TabIndex = 1
        Me.picInv.Location = New System.Drawing.Point(609, 147)
        Me.picInv.Size = New System.Drawing.Size(162, 162)
        Me.picInv.CausesValidation = False
        Me.picInv.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.picInv.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(0, Byte))
        Me.picInv.Font = New System.Drawing.Font("MS Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'cmdMenu_0
        '
        Me.cmdMenu_0.Name = "cmdMenu_0"
        Me.cmdMenu_0.Location = New System.Drawing.Point(625, 133)
        Me.cmdMenu_0.Size = New System.Drawing.Size(127, 30)
        Me.cmdMenu_0.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.cmdMenu_0.BackColor = System.Drawing.SystemColors.Control
        '
        'cmdMenu_4
        '
        Me.cmdMenu_4.Name = "cmdMenu_4"
        Me.cmdMenu_4.Location = New System.Drawing.Point(625, 291)
        Me.cmdMenu_4.Size = New System.Drawing.Size(127, 30)
        Me.cmdMenu_4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.cmdMenu_4.BackColor = System.Drawing.SystemColors.Control
        '
        'cmdMen
        '
        Me.cmdMen.Name = "cmdMen"
        Me.cmdMen.Location = New System.Drawing.Point(724, 81)
        Me.cmdMen.Size = New System.Drawing.Size(73, 36)
        Me.cmdMen.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.cmdMen.BackColor = System.Drawing.SystemColors.Control
        '
        'nomodoseguro
        '
        Me.nomodoseguro.Name = "nomodoseguro"
        Me.nomodoseguro.Visible = False
        Me.ToolTip1.SetToolTip(Me.nomodoseguro, "Seguro")
        Me.nomodoseguro.Location = New System.Drawing.Point(623, 514)
        Me.nomodoseguro.Size = New System.Drawing.Size(20, 17)
        Me.nomodoseguro.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.nomodoseguro.BackColor = System.Drawing.SystemColors.Control
        Me.nomodoseguro.Image = CType(Resources.GetObject("nomodoseguro.Image"), System.Drawing.Bitmap)
        '
        'cmdInv
        '
        Me.cmdInv.Name = "cmdInv"
        Me.cmdInv.Location = New System.Drawing.Point(578, 81)
        Me.cmdInv.Size = New System.Drawing.Size(73, 36)
        Me.cmdInv.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.cmdInv.BackColor = System.Drawing.SystemColors.Control
        '
        'Hpshp
        '
        Me.Hpshp.Name = "Hpshp"
        Me.Hpshp.Location = New System.Drawing.Point(590, 396)
        Me.Hpshp.Size = New System.Drawing.Size(92, 9)
        Me.Hpshp.Text = ""
        Me.Hpshp.BackColor = System.Drawing.Color.FromArgb(CType(192, Byte), CType(0, Byte), CType(0, Byte))
        '
        'AGUAsp
        '
        Me.AGUAsp.Name = "AGUAsp"
        Me.AGUAsp.Location = New System.Drawing.Point(696, 451)
        Me.AGUAsp.Size = New System.Drawing.Size(92, 9)
        Me.AGUAsp.Text = ""
        Me.AGUAsp.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(192, Byte))
        '
        'COMIDAsp
        '
        Me.COMIDAsp.Name = "COMIDAsp"
        Me.COMIDAsp.Location = New System.Drawing.Point(696, 424)
        Me.COMIDAsp.Size = New System.Drawing.Size(92, 9)
        Me.COMIDAsp.Text = ""
        Me.COMIDAsp.BackColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(192, Byte), CType(0, Byte))
        '
        'lblHAM
        '
        Me.lblHAM.Name = "lblHAM"
        Me.lblHAM.TabIndex = 16
        Me.lblHAM.Location = New System.Drawing.Point(696, 422)
        Me.lblHAM.Size = New System.Drawing.Size(92, 12)
        Me.lblHAM.Text = "0/0"
        Me.lblHAM.BackColor = System.Drawing.Color.Transparent
        Me.lblHAM.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(255, Byte))
        Me.lblHAM.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblHAM.Font = New System.Drawing.Font("Tahoma", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'lblAG
        '
        Me.lblAG.Name = "lblAG"
        Me.lblAG.TabIndex = 13
        Me.lblAG.Location = New System.Drawing.Point(629, 576)
        Me.lblAG.Size = New System.Drawing.Size(23, 13)
        Me.lblAG.Text = "0"
        Me.lblAG.BackColor = System.Drawing.Color.Transparent
        Me.lblAG.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(255, Byte))
        '
        'lblMP
        '
        Me.lblMP.Name = "lblMP"
        Me.lblMP.TabIndex = 11
        Me.lblMP.Location = New System.Drawing.Point(590, 422)
        Me.lblMP.Size = New System.Drawing.Size(92, 12)
        Me.lblMP.Text = "0/0"
        Me.lblMP.BackColor = System.Drawing.Color.Transparent
        Me.lblMP.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(255, Byte))
        Me.lblMP.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblMP.Font = New System.Drawing.Font("Tahoma", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'lblExp
        '
        Me.lblExp.Name = "lblExp"
        Me.lblExp.TabIndex = 6
        Me.lblExp.Location = New System.Drawing.Point(599, 59)
        Me.lblExp.Size = New System.Drawing.Size(122, 12)
        Me.lblExp.Text = "0% (0/0)"
        Me.lblExp.BackColor = System.Drawing.Color.Transparent
        Me.lblExp.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(0, Byte))
        Me.lblExp.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.lblExp.Font = New System.Drawing.Font("Tahoma", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'GldLbl
        '
        Me.GldLbl.Name = "GldLbl"
        Me.GldLbl.TabIndex = 4
        Me.GldLbl.Location = New System.Drawing.Point(716, 387)
        Me.GldLbl.Size = New System.Drawing.Size(75, 12)
        Me.GldLbl.Text = "0"
        Me.GldLbl.BackColor = System.Drawing.Color.Transparent
        Me.GldLbl.ForeColor = System.Drawing.Color.FromArgb(CType(255, Byte), CType(255, Byte), CType(255, Byte))
        Me.GldLbl.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.GldLbl.Font = New System.Drawing.Font("Tahoma", 6.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        '
        'lblInvInfo
        '
        Me.lblInvInfo.Name = "lblInvInfo"
        Me.lblInvInfo.TabIndex = 8
        Me.lblInvInfo.Location = New System.Drawing.Point(607, 316)
        Me.lblInvInfo.Size = New System.Drawing.Size(163, 41)
        Me.lblInvInfo.Text = ""
        Me.lblInvInfo.BackColor = System.Drawing.Color.Transparent
        Me.lblInvInfo.ForeColor = System.Drawing.SystemColors.HighlightText
        Me.lblInvInfo.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'frmMain
        '
        Me.ClientSize = New System.Drawing.Size(809, 607)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Command1, Me.picMacro_0, Me.picMacro_9, Me.Minimap, Me.RecChat, Me.picInv, Me.cmdMenu_0, Me.cmdMenu_4, Me.cmdMen, Me.nomodoseguro, Me.cmdInv, Me.Hpshp, Me.AGUAsp, Me.COMIDAsp, Me.lblHAM, Me.lblAG, Me.lblMP, Me.lblExp, Me.GldLbl, Me.lblInvInfo})
        Me.Name = "frmMain"
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(64, Byte), CType(64, Byte), CType(64, Byte))
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(0, Byte), CType(0, Byte), CType(0, Byte))
        Me.KeyPreview = True
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.Icon = CType(Resources.GetObject("frmMain.Icon"), System.Drawing.Icon)
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Inmortal AO"
        Me.Visible = False
        Me.picMacro.SetIndex(picMacro_0, CType(0, Short))
        Me.picMacro.SetIndex(picMacro_1, CType(1, Short))
        Me.picMacro.SetIndex(picMacro_3, CType(3, Short))
        Me.picMacro.SetIndex(picMacro_2, CType(2, Short))
        Me.picMacro.SetIndex(picMacro_4, CType(4, Short))
        Me.picMacro.SetIndex(picMacro_5, CType(5, Short))
        Me.picMacro.SetIndex(picMacro_6, CType(6, Short))
        Me.picMacro.SetIndex(picMacro_7, CType(7, Short))
        Me.picMacro.SetIndex(picMacro_8, CType(8, Short))
        Me.picMacro.SetIndex(picMacro_9, CType(9, Short))
        Me.picMacro.SetIndex(picMacro_10, CType(10, Short))
        Me.cmdMenu.SetIndex(cmdMenu_5, CType(5, Short))
        Me.cmdMenu.SetIndex(cmdMenu_3, CType(3, Short))
        Me.cmdMenu.SetIndex(cmdMenu_2, CType(2, Short))
        Me.cmdMenu.SetIndex(cmdMenu_1, CType(1, Short))
        Me.cmdMenu.SetIndex(cmdMenu_0, CType(0, Short))
        Me.cmdMenu.SetIndex(cmdMenu_4, CType(4, Short))
        'Me.Label2.SetIndex(Label2_0, CType(0, Short))
        'Me.cmdMoverHechi.SetIndex(cmdMoverHechi_0, CType(0, Short))
        'Me.cmdMoverHechi.SetIndex(cmdMoverHechi_1, CType(1, Short))
        CType(Me.picMacro, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdMenu, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.Label2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.cmdMoverHechi, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Minimap.ResumeLayout(False)
        Me.RecChat.ResumeLayout(False)
        CType(Me.RecChat, System.ComponentModel.ISupportInitialize).EndInit()
        Me.picInv.ResumeLayout(False)
        Me.lblHAM.ResumeLayout(False)
        Me.lblMP.ResumeLayout(False)
        Me.lblExp.ResumeLayout(False)
        Me.GldLbl.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
#Const IndexCurFile = 1
#If CompileAll Or CompileAllFRM Or ((IndexCurFile >= CompileFilesIndexMin) And (IndexCurFile <= CompileFilesIndexMax)) Then
#Const CompileAll_frmMain = True
#Else
'#Const CompileAll_frmMain = True
#End If

End Class