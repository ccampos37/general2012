VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmProcesolista 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   390
      TabIndex        =   2
      Top             =   60
      Width           =   6615
      Begin VB.CommandButton cProcesa 
         Caption         =   "Procesa"
         Height          =   315
         Left            =   5370
         TabIndex        =   6
         Top             =   180
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Externo"
         Height          =   195
         Index           =   1
         Left            =   4200
         TabIndex        =   4
         Top             =   240
         Width           =   1125
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Propio"
         Height          =   195
         Index           =   0
         Left            =   3150
         TabIndex        =   3
         Top             =   210
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "PROCESA CON  INFORMACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   150
         TabIndex        =   5
         Top             =   240
         Width           =   2745
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   600
      Left            =   7290
      Top             =   1050
   End
   Begin MSComctlLib.ProgressBar Bar1 
      Height          =   345
      Left            =   330
      TabIndex        =   0
      Top             =   1080
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   609
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label Label1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   690
      Width           =   6675
   End
End
Attribute VB_Name = "FrmProcesolista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sqlini, sqlfin As String
Dim snprecio, sdventa, sdreparto As Double
Dim stipo As Integer
Dim wflag As Integer

Private Sub cProcesa_Click()
    If stipo = 1 Then
       Procesadata
    ElseIf stipo = 2 Then
       Procesagrilla
    End If
    
End Sub

Private Sub Form_Load()
   MostrarForm Me, "C2"
   Me.Show
   Me.Caption = "Procesando Informacion...."
  ' Me.Show
  Call cProcesa_Click
  DoEvents

End Sub




Public Function Procesadata()
   Dim rori As New ADODB.Recordset
   Dim acmd As New ADODB.Command
   Dim nsql As String
   On Error Resume Next
      
   wflag = 1
   
  
   If Option1(1).Value Then
        Set acmd.ActiveConnection = VGgeneral
        acmd.CommandText = "vt_listaproducto_pro"
        acmd.CommandType = adCmdStoredProc
        acmd.CommandTimeout = 0
        acmd.Prepared = True
        acmd.Parameters("@base") = VGcnx.DefaultDatabase
        acmd.Parameters("@almacen") = "%"
        Set rori = acmd.Execute
        Set acmd = Nothing
   Else
        Set rori = VGcnx.Execute(sqlini)
   End If
   If rori.RecordCount > 0 Then
      Bar1.Value = 0
      Bar1.Max = rori.RecordCount + 1
      rori.MoveFirst
      Do Until rori.EOF
         nsql = "INSERT INTO " & sqlfin & _
                "(productocodigo,productodescripcion,productoprecvta," & _
                "productodescrcorta, grupovtacodigo,productofamiliacodigo," & _
                "productocategoriacodigo,productotipo,productoporcimpto," & _
                "productoestunidreferencia,unidadfactorconv," & _
                "productoprecvtaofi,productoprecvtareparto,monedacodigo,unidadcodigo,unidadreferencial,almacencodigo)" & _
                " VALUES (" & _
                "'" & rori!productocodigo & "','" & rori!productodescripcion & "'," & _
                IIf(IsNull(rori!productoprecvta), CDbl(0), rori!productoprecvta) & "," & _
                "'" & Escadena(rori!productodescrcorta) & "','" & Escadena(rori!grupovtacodigo) & "','" & Escadena(rori!productofamiliacodigo) & "'," & _
                "'" & Escadena(rori!productocategoriacodigo) & "','" & Escadena(rori!productotipo) & "'," & IIf(IsNull(rori!productoporcimpto), CDbl(0), rori!productoporcimpto) & "," & _
                IIf(rori!productoestunidreferencia, "1", "0") & "," & IIf(IsNull(rori!unidadfactorconv), CDbl(0), rori!unidadfactorconv) & "," & _
                IIf(IsNull(rori!productoprecvta), CDbl(0), rori!productoprecvta) * (1 + (sdventa / 100)) & "," & _
                IIf(IsNull(rori!productoprecvta), CDbl(0), rori!productoprecvta) * (1 + (sdreparto / 100)) & "," & _
                 "'" & Escadena(rori!monedacodigo) & "','" & Escadena(rori!unidadcodigo) & "','" & Escadena(rori!unidadreferencial) & "','" & Escadena(rori!almacencodigo) & "')"
         VGcnx.Execute nsql
         
         Bar1.Value = Bar1.Value + 1
         DoEvents
         rori.MoveNext
      Loop
   End If
   Bar1.Value = Bar1.Max
   Set rori = Nothing
   DoEvents
   wflag = 0
   Unload FrmProcesolista
End Function

Public Function Procesagrilla()
   Dim nsql As String
   Dim adll As New dllgeneral.dll_general
   
    On Error GoTo nerror
    wflag = 1
    Bar1.Value = 0
    Bar1.Max = FrmListaPrecios.TDBGrid2.ApproxCount + 1
    VGcnx.BeginTrans
    FrmListaPrecios.TDBGrid2.MoveLast
    FrmListaPrecios.TDBGrid2.MoveFirst
    Do Until FrmListaPrecios.TDBGrid2.EOF
       'If adll.VerificaDatoExistente(VGcnx, "select * from " & FrmListaPrecios.TDBGrid1.Columns(1).Text & " where productocodigo='" & FrmListaPrecios.TDBGrid2.Columns(0).Text & "'") = 0 Then
'      If adll.VerificaDatoExistente(VGcnx, "select * from " & sqlfin & " where productocodigo='" & FrmListaPrecios.TDBGrid2.Columns(0).Text & "' and almacencodigo='" & FrmListaPrecios.TDBGrid2.Columns(3).Text & "'") = 0 Then
        '  nsql = "INSERT INTO " & FrmListaPrecios.TDBGrid1.Columns(1).Text & _
        '          "Select * from vt_producto where productocodigo='" & FrmListaPrecios.TDBGrid2.Columns(0).Text & "'"
        ''  nsql = "INSERT INTO " & sqlfin & _
        ''          "Select * from vt_producto where productocodigo='" & FrmListaPrecios.TDBGrid2.Columns(0).Text & "'"

'      End If
       
        nsql = "UPDATE " & sqlfin & _
                " Set productocodigo='" & FrmListaPrecios.TDBGrid2.Columns(0).Text & "'," & _
                "    productodescripcion='" & FrmListaPrecios.TDBGrid2.Columns(1).Text & "'," & _
                "    productoprecvta=" & CDbl(FrmListaPrecios.TDBGrid2.Columns(2).Text) & _
                " Where productocodigo='" & FrmListaPrecios.TDBGrid2.Columns(0).Text & "' and almacencodigo='" & FrmListaPrecios.PTalmacen & "'"
                      
       VGcnx.Execute nsql
       
       Bar1.Value = Bar1.Value + 1
       DoEvents
       FrmListaPrecios.TDBGrid2.MoveNext
    Loop
   Bar1.Value = Bar1.Max
   VGcnx.CommitTrans
   DoEvents
   wflag = 0
   'Unload FrmProcesolista
   
nerror:
  If Err Then
    'Err = 0
    MsgBox Err.Number & "-" & Err.Description
    'If VGcnx.State = adStateExecuting Then
    'VGcnx.RollbackTrans
    'End If
    DoEvents
    Exit Function
    Resume
  End If
  
End Function

Private Sub Timer1_Timer()
  If Len(Trim(Label1)) = 0 Or Left(Label1, 3) = "Pro" Then
     Label1.Caption = "Espere un Momento...!!!"
  Else
     Label1.Caption = "Procesando Informacion...!!!"
  End If
End Sub



Public Property Let bsqlini(pdata As String)
   sqlini = pdata
End Property


Public Property Let bsqlfin(pdata As String)
   sqlfin = pdata
End Property

Public Property Let bfactorprecio(pdata As String)
   snprecio = pdata
End Property

Public Property Let bfactordvta(pdata As String)
   sdventa = pdata
End Property


Public Property Let bfactordreparto(pdata As String)
   sdreparto = pdata
End Property

Public Property Let Btipo(pdata As Integer)
   stipo = pdata
End Property

