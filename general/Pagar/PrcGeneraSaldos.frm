VERSION 5.00
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form PrcGeneraSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Regeneracion de Saldos"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6345
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "Tipo Proceso"
      Height          =   1005
      Left            =   300
      TabIndex        =   12
      Top             =   240
      Width           =   5565
      Begin VB.CommandButton cProcesa 
         Caption         =   "&Proceso"
         Height          =   375
         Left            =   4260
         TabIndex        =   16
         Top             =   180
         Width           =   1155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "General"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   13
         Top             =   330
         Value           =   -1  'True
         Width           =   1185
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   330
         Left            =   1320
         TabIndex        =   15
         Top             =   600
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   582
         XcodMaxLongitud =   11
         xcodwith        =   800
         NomTabla        =   "cp_proveedor"
         TituloAyuda     =   "Ayuda de Clientes"
         ListaCampos     =   "clientecodigo(1),clienterazonsocial(1)"
         XcodCampo       =   "clientecodigo"
         XListCampo      =   "clienterazonsocial"
         ListaCamposDescrip=   "Codigo,Descripcion"
         ListaCamposText =   "clientecodigo,clienterazonsocial"
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Individual"
         Height          =   195
         Index           =   1
         Left            =   210
         TabIndex        =   14
         Top             =   660
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   270
      TabIndex        =   0
      Top             =   3540
      Width           =   5655
      Begin VB.Frame Frame5 
         Height          =   915
         Left            =   150
         TabIndex        =   9
         Top             =   240
         Width           =   5355
         Begin VB.Label Label5 
            Height          =   195
            Left            =   180
            TabIndex        =   11
            Top             =   540
            Width           =   4965
         End
         Begin VB.Label Label4 
            Caption         =   "Documento Abonado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1740
            TabIndex        =   10
            Top             =   225
            Width           =   1800
         End
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   270
      TabIndex        =   3
      Top             =   1350
      Width           =   5625
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   3030
         TabIndex        =   7
         Top             =   600
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Registros a Procesar"
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
         Index           =   1
         Left            =   960
         TabIndex        =   6
         Top             =   600
         Width           =   1905
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   3030
         TabIndex        =   5
         Top             =   210
         Width           =   1245
      End
      Begin VB.Label Label1 
         Caption         =   "Registros Procesados"
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
         Index           =   0
         Left            =   960
         TabIndex        =   4
         Top             =   240
         Width           =   1905
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1005
      Left            =   270
      TabIndex        =   1
      Top             =   2430
      Width           =   5655
      Begin MSComctlLib.ProgressBar Bar1 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   540
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "PrcGeneraSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nopcion As String
Dim SQL As String

Private Sub cProcesa_Click()
    Call ProcesaSaldos
End Sub

Private Sub Form_Load()
    MostrarForm Me, "C2"
    DoEvents
    Ctr_Ayuda1.conexion VGCNx
End Sub

Public Sub ProcesaSaldos()
    Dim rb As New ADODB.Recordset
    Dim rbabono As New ADODB.Recordset
    Dim rbcargo As New ADODB.Recordset
    Dim ncuenta As Double
    Dim ntotal As Double
        
    On Error GoTo nerror
    
    If Option1(0).Value Then
        nopcion = "%%"
    Else
        nopcion = Trim$(Escadena(Ctr_Ayuda1.xclave))
    End If
    
    If Len(Trim$(nopcion)) = 0 Then
        MsgBox "No existen datos a consultar....Verifique!!!", vbInformation, MsgTitle
        Exit Sub
    End If
    
    ' Colocar los saldos de los documentos sin pagos
    ' cargoapeimppag=0 : cargoapeflgcan=0:cargoapefeccan=null
    VGCNx.BeginTrans
    SQL = "Update cp_cargo Set cargoapeimppag=0,cargoapeflgcan=0,cargoapefeccan=null"
    If Ctr_Ayuda1.xclave <> "" Then SQL = SQL & " where clientecodigo='" & Ctr_Ayuda1.xclave & "'"
    VGCNx.Execute (SQL)
    
    DoEvents
    ' Acumula los abonos de los documentos
    Set rb = VGCNx.Execute("select distinct abonocancli,documentoabono,abononumdoc,empresacodigo" & _
                        " From cp_abono " & _
                        " Inner join cp_tipodocumento " & _
                        " on cp_abono.documentoabono=cp_tipodocumento.tdocumentocodigo " & _
                        " where abonocancli like '" & nopcion & "'" & _
                        " and isnull(abonocanflreg,0)<>1 ")
                        
VGCNx.CommitTrans
If rb.RecordCount > 0 Then
   rb.MoveLast
   ncuenta = 0
   ntotal = rb.RecordCount
   Label2(1) = Numero(ntotal)
   Bar1.Max = ntotal + 1
   rb.MoveFirst
   Do Until rb.EOF
      ncuenta = ncuenta + 1
      Label3 = Numero((ncuenta / ntotal) * 100)
      Bar1.Value = Bar1.Value + 1
      Label2(0) = Numero(ncuenta)
        
     'Actualizar los saldos en cp_cargo
      SQL = "select * from cp_cargo where empresacodigo='" & rb!empresacodigo & "' and documentocargo='" & rb.Fields(1) & "'"
      SQL = SQL & " and cargonumdoc='" & rb.Fields(2) & "' and clientecodigo ='" & rb.Fields(0) & "'"
      Set rbcargo = VGCNx.Execute(SQL)
      If rbcargo.RecordCount > 0 Then
         If rbcargo.Fields("monedacodigo") = g_TipoSol Then
            Set rbabono = VGCNx.Execute("select documentoabono,abononumdoc," & _
                " round(sum( case abonocanmoncan when '02' then " & _
                " (isnull(abonocanimpcan,0)*isnull(abonocantipcam,1)) else 0 end),2)," & _
                " round(sum( case abonocanmoncan when '01' then " & _
                " isnull(abonocanimpcan,0) else 0 end),2) " & _
                " From cp_abono Inner join cp_tipodocumento " & _
                " on cp_abono.documentoabono=cp_tipodocumento.tdocumentocodigo " & _
                " where empresacodigo='" & rbcargo!empresacodigo & "' and cp_abono.documentoabono='" & rbcargo!documentocargo & "' and cp_abono.abononumdoc='" & rbcargo!cargonumdoc & "' and abonocancli ='" & rbcargo!clientecodigo & "' " & _
                " and isnull(abonocanflreg,0)<>1 group by empresacodigo,abonocancli,documentoabono,abononumdoc")
          ElseIf rbcargo.Fields("monedacodigo") = g_TipoDolar Then
                 Set rbabono = VGCNx.Execute("select documentoabono,abononumdoc," & _
                     " round(sum( case abonocanmoncan when '02' then " & _
                     " isnull(abonocanimpcan,0) else 0 end),2)," & _
                     " round(sum( case abonocanmoncan when '01' then " & _
                     " (isnull(abonocanimpsol,0)) else 0 end),2) " & _
                     " From cp_abono Inner join cp_tipodocumento " & _
                     " on cp_abono.documentoabono=cp_tipodocumento.tdocumentocodigo " & _
                     " where empresacodigo='" & rbcargo!empresacodigo & "' and cp_abono.documentoabono='" & rbcargo!documentocargo & "' and cp_abono.abononumdoc='" & rbcargo!cargonumdoc & "' and abonocancli like '" & rbcargo.Fields("clientecodigo") & "' " & _
                     " and isnull(abonocanflreg,0)<>1 group by empresacodigo,abonocancli,documentoabono,abononumdoc")
                  
         End If
         If rbabono.RecordCount > 0 Then
            VGCNx.Execute "Update cp_cargo " & _
                  "Set cargoapeimppag=" & rbabono.Fields(2) + rbabono.Fields(3) & _
                  " Where empresacodigo='" & rbcargo!empresacodigo & "' and documentocargo='" & rbcargo.Fields("documentocargo") & "' and cargonumdoc='" & rbcargo.Fields("cargonumdoc") & "' and clientecodigo like '" & rbcargo.Fields("clientecodigo") & "'"
            SQL = "Update cp_cargo Set cargoapeflgcan= case Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) when 0 then '1' else '0' end "
'            SQL = SQL & " cargoapefeccan=case Round(isnull(cargoapeimpape,0),2)-Round(isnull(cargoapeimppag,0),2) when 0 then '" & rbabono.Fields(2) & "' else null end  "
            SQL = SQL & " Where empresacodigo='" & rbcargo!empresacodigo & "' and documentocargo='" & rbcargo.Fields("documentocargo") & "' and cargonumdoc='" & rbcargo.Fields("cargonumdoc") & "' and clientecodigo like '" & rbcargo.Fields("clientecodigo") & "'"
            VGCNx.Execute (SQL)
         End If
     End If
     rbcargo.Close
     Set rbcargo = Nothing
 '    Label5 = rb.Fields(0) & "-" & rb.Fields(1)
     DoEvents
     rb.MoveNext
  Loop
  Bar1.Value = Bar1.Max
End If
rb.Close
Set rb = Nothing
                
MsgBox "Proceso Terminado Satisfactoriamente ...!!!", vbInformation, MsgTitle
    

nerror:
    If Err Then
        Err = 0
        MsgBox "El Proceso no se culmino ...!!!" & Chr(13) & Chr(10) & Err.Number & "-" & Err.Description, vbInformation, MsgTitle
        Exit Sub
        Resume
    End If

End Sub


