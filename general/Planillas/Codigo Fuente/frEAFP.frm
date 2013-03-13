VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frEAFP 
   Caption         =   "Edición de Fondos de Pensiones"
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4800
   Icon            =   "frEAFP.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   4800
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2513
      TabIndex        =   13
      Top             =   2940
      Width           =   1185
   End
   Begin VB.CommandButton cmAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   1103
      TabIndex        =   12
      Top             =   2940
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Información General"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      Begin AplisetControlText.Aplitext xRem 
         Height          =   285
         Left            =   2070
         TabIndex        =   11
         Top             =   2130
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Text            =   "0.00"
         Redondear       =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xSeg 
         Height          =   285
         Left            =   2070
         TabIndex        =   10
         Top             =   1422
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Text            =   "0.00"
         Redondear       =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xApor 
         Height          =   285
         Left            =   2070
         TabIndex        =   9
         Top             =   1068
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Text            =   "0.00"
         Redondear       =   -1  'True
         TipoDato        =   "N"
      End
      Begin AplisetControlText.Aplitext xNom 
         Height          =   285
         Left            =   2070
         TabIndex        =   8
         Top             =   714
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   503
         MaxLength       =   25
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCod 
         Height          =   285
         Left            =   2070
         TabIndex        =   7
         Top             =   360
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         MaxLength       =   2
         Text            =   ""
         TipoCodigo      =   -1  'True
      End
      Begin AplisetControlText.Aplitext xTope 
         Height          =   285
         Left            =   2070
         TabIndex        =   14
         Top             =   1800
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   503
         Text            =   "0"
         TipoDato        =   "N"
      End
      Begin VB.Label Label6 
         Caption         =   "Remun. Variable"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   2190
         Width           =   1515
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tope de Seguro"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   1824
         Width           =   1155
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Seguro (%)"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   1458
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aportación Obligatoria"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   1092
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   726
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   180
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
   End
End
Attribute VB_Name = "frEAFP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CMACEPTAR_CLICK()
    If VPTAREA = "NUEVO" And frAFPs.EXISTE(xCod.Text) Then
        MsgBox "EL CÓDIGO INGRESADO YA EXISTE, POR FAVOR INTENTE DE NUEVO", vbCritical
        xCod.SetFocus
        Exit Sub
    End If
    If xNom.Text = "" Then
        MsgBox "DEBE INGRESAR UN NOMBRE DE AFP VÁLIDO", vbCritical
        xNom.SetFocus
        Exit Sub
    End If
    If VPTAREA = "NUEVO" Then
        DBSYSTEM.Execute "INSERT INTO AFPS (CODAFP,NOMBRE,APOROBLI,SEGURO,TOPESEGURO,COMISIONRA) SELECT '" & xCod.Text & "','" & xNom.Text & "'," & xApor.Text & "," & xSeg.Text & "," & xTope.Text & "," & xRem.Text
    Else
        'EDITAR
        DBSYSTEM.Execute "UPDATE AFPS SET NOMBRE='" & xNom.Text & "',APOROBLI=" & xApor.Text & ",SEGURO=" & xSeg.Text & ",TOPESEGURO=" & xTope.Text & ",COMISIONRA=" & xRem.Text & " WHERE CODAFP='" & xCod.Text & "'"
    End If
    Unload Me
End Sub

Private Sub CMCANCELAR_CLICK()

    Unload Me
End Sub

Private Sub Form_Activate()
    If VPTAREA = "EDITAR" Then
        xCod.Locked = True
        xNom.SetFocus
    End If
End Sub

Private Sub Form_Load()
    If VPTAREA = "EDITAR" Then
        With frAFPs.lvAFPs.SelectedItem
            xCod.Text = .Text
            xNom.Text = .SubItems(1)
            xApor.Text = .SubItems(2)
            xSeg.Text = .SubItems(3)
            xTope.Text = .SubItems(4)
            xRem.Text = .SubItems(5)
        End With
    End If
End Sub

