Public Class FrmCierres
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
    Friend WithEvents Command2 As System.Windows.Forms.Button
    Friend WithEvents Command1 As System.Windows.Forms.Button
    Friend WithEvents Frame1 As System.Windows.Forms.GroupBox
    Friend TxFeranno As AxTextFer.AxTxFer
    Friend TxFermes As AxTextFer.AxTxFer
    Friend WithEvents Label2 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmCierres))
        Me.Command2 = New System.Windows.Forms.Button()
        Me.Command1 = New System.Windows.Forms.Button()
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.TxFeranno = New AxTextFer.AxTxFer()
        Me.TxFermes = New AxTextFer.AxTxFer()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Command2
        '
        Me.Command2.Name = "Command2"
        Me.Command2.TabIndex = 6
        Me.Command2.Location = New System.Drawing.Point(194, 138)
        Me.Command2.Size = New System.Drawing.Size(82, 66)
        Me.Command2.Text = "Salir"
        Me.Command2.BackColor = System.Drawing.SystemColors.Control
        Me.Command2.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Command1
        '
        Me.Command1.Name = "Command1"
        Me.Command1.TabIndex = 5
        Me.Command1.Location = New System.Drawing.Point(57, 138)
        Me.Command1.Size = New System.Drawing.Size(82, 66)
        Me.Command1.Text = "Aceptar"
        Me.Command1.BackColor = System.Drawing.SystemColors.Control
        Me.Command1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'Frame1
        '
        Me.Frame1.Controls.AddRange(New System.Windows.Forms.Control() {Me.TxFeranno, Me.TxFermes, Me.Label2})
        Me.Frame1.Name = "Frame1"
        Me.Frame1.TabIndex = 0
        Me.Frame1.Location = New System.Drawing.Point(40, 24)
        Me.Frame1.Size = New System.Drawing.Size(260, 98)
        Me.Frame1.Text = ""
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'TxFeranno
        '
        Me.TxFeranno.Name = "TxFeranno"
        Me.TxFeranno.TabIndex = 3
        Me.TxFeranno.Font = New System.Drawing.Font("MS Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TxFeranno.Location = New System.Drawing.Point(129, 24)
        Me.TxFeranno.Size = New System.Drawing.Size(66, 25)
        '
        'TxFermes
        '
        Me.TxFermes.Name = "TxFermes"
        Me.TxFermes.TabIndex = 4
        Me.TxFermes.Font = New System.Drawing.Font("MS Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.TxFermes.Location = New System.Drawing.Point(129, 57)
        Me.TxFermes.Size = New System.Drawing.Size(66, 25)
        '
        'Label2
        '
        Me.Label2.Name = "Label2"
        Me.Label2.TabIndex = 2
        Me.Label2.Font = New System.Drawing.Font("MS Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(204, Byte))
        Me.Label2.Location = New System.Drawing.Point(16, 65)
        Me.Label2.Size = New System.Drawing.Size(82, 25)
        Me.Label2.Text = "Mes de cierre"
        Me.Label2.BackColor = System.Drawing.SystemColors.Control
        Me.Label2.ForeColor = System.Drawing.SystemColors.ControlText
        '
        'FrmCierres
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(359, 237)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Command2, Me.Command1, Me.Frame1})
        Me.Name = "FrmCierres"
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.MinimizeBox = True
        Me.MaximizeBox = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.WindowsDefaultLocation
        Me.Text = "Control de cierres"
        Me.TxFeranno.ResumeLayout(False)
        Me.TxFermes.ResumeLayout(False)
        Me.Label2.ResumeLayout(False)
        Me.Frame1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

	'=========================================================
    Private Sub Command1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command1.Click
'#Const def_Command1_Click = True
#If def_Command1_Click
        ' VBto Upgrade Warning: rrsql As ADODB.Recordset	OnWrite(Integer)
        Dim rrsql As ADODB.Recordset
        rrsql = VGCNx.Execute(" update cs_sistema set mesdecierre='" & TxFeranno.valor & TxFermes.valor & "'")
        rrsql = Nothing
#End If	' def_Command1_Click
    End Sub

    Private Sub Command2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Command2.Click
'#Const def_Command2_Click = True
#If def_Command2_Click
        Close()
#End If	' def_Command2_Click
    End Sub

    Private Sub FrmCierres_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
'#Const def_Form_Load = True
#If def_Form_Load
        TxFeranno.valor = VGParametros.mesdecierre.Substring(0, 4)
        TxFermes.valor = VGParametros.mesdecierre.Substring(VGParametros.mesdecierre.Length-2, 2)
#End If	' def_Form_Load
    End Sub

End Class