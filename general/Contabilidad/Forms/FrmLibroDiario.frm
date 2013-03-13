VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form FrmLibroDiario 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Libro Diario"
   ClientHeight    =   5175
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6165
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6165
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraResumido 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Diario General Resumido"
      Height          =   1815
      Left            =   480
      TabIndex        =   17
      Top             =   2235
      Width           =   5145
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   330
         Index           =   0
         Left            =   1275
         TabIndex        =   18
         Top             =   435
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         XcodMaxLongitud =   3
         xcodwith        =   600
         NomTabla        =   "ct_asiento"
         ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
         XcodCampo       =   "asientocodigo"
         XListCampo      =   "asientodescripcion"
         ListaCamposDescrip=   "Código,Descripcion"
         ListaCamposText =   "asientocodigo,asientodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda1 
         Height          =   330
         Index           =   1
         Left            =   1275
         TabIndex        =   19
         Top             =   750
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         XcodMaxLongitud =   4
         xcodwith        =   600
         NomTabla        =   "ct_subasiento"
         ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1)"
         XcodCampo       =   "subasientocodigo"
         XListCampo      =   "subasientodescripcion"
         ListaCamposDescrip=   "Código,Descripcion"
         ListaCamposText =   "subasientocodigo,subasientodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Asiento"
         Height          =   240
         Left            =   180
         TabIndex        =   21
         Top             =   465
         Width           =   885
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "SubAsiento"
         Height          =   300
         Left            =   180
         TabIndex        =   20
         Top             =   795
         Width           =   960
      End
   End
   Begin VB.Frame fraDetallado 
      Caption         =   "Diario General Detallado"
      Height          =   1815
      Left            =   480
      TabIndex        =   8
      Top             =   2235
      Width           =   5145
      Begin MSComCtl2.DTPicker DTPickerFecFinal 
         Height          =   300
         Left            =   1290
         TabIndex        =   9
         Top             =   555
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   529
         _Version        =   393216
         Format          =   17170433
         CurrentDate     =   37474
      End
      Begin MSComCtl2.DTPicker DTPickerFecInicio 
         Height          =   300
         Left            =   1305
         TabIndex        =   10
         Top             =   225
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   529
         _Version        =   393216
         CalendarBackColor=   14876158
         Format          =   17170433
         CurrentDate     =   37474
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   330
         Index           =   0
         Left            =   1275
         TabIndex        =   11
         Top             =   960
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         XcodMaxLongitud =   3
         xcodwith        =   600
         NomTabla        =   "ct_asiento"
         ListaCampos     =   "asientocodigo(1),asientodescripcion(1)"
         XcodCampo       =   "asientocodigo"
         XListCampo      =   "asientodescripcion"
         ListaCamposDescrip=   "Código,Descripcion"
         ListaCamposText =   "asientocodigo,asientodescripcion"
         Requerido       =   0   'False
      End
      Begin ctrlayuda_f.Ctr_Ayuda Ctr_Ayuda2 
         Height          =   330
         Index           =   1
         Left            =   1275
         TabIndex        =   12
         Top             =   1275
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   582
         XcodMaxLongitud =   4
         xcodwith        =   600
         NomTabla        =   "ct_subasiento"
         ListaCampos     =   "subasientocodigo(1),subasientodescripcion(1)"
         XcodCampo       =   "subasientocodigo"
         XListCampo      =   "subasientodescripcion"
         ListaCamposDescrip=   "Código,Descripcion"
         ListaCamposText =   "subasientocodigo,subasientodescripcion"
         Requerido       =   0   'False
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   300
         Left            =   195
         TabIndex        =   16
         Top             =   300
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   285
         Left            =   195
         TabIndex        =   15
         Top             =   615
         Width           =   1065
      End
      Begin VB.Label Label5 
         Caption         =   "Asiento"
         Height          =   225
         Left            =   195
         TabIndex        =   14
         Top             =   1005
         Width           =   1035
      End
      Begin VB.Label Label6 
         Caption         =   "SubAsiento"
         Height          =   270
         Left            =   180
         TabIndex        =   13
         Top             =   1335
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Height          =   1260
      Left            =   480
      TabIndex        =   4
      Top             =   120
      Width           =   5130
      Begin VB.OptionButton optOpcion 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Diario General Resumido Tipo 1"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   2655
      End
      Begin VB.OptionButton optOpcion 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Diario General Detallado"
         Height          =   375
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   2295
      End
      Begin VB.OptionButton optOpcion 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Diario General Resumido Tipo 2"
         Height          =   375
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Cancelar"
      Height          =   345
      Index           =   1
      Left            =   3255
      TabIndex        =   3
      Top             =   4440
      Width           =   1230
   End
   Begin VB.CommandButton cmdBotones 
      Caption         =   "&Aceptar"
      Height          =   345
      Index           =   0
      Left            =   1650
      TabIndex        =   2
      Top             =   4440
      Width           =   1230
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Seleccionar Mes"
      Height          =   705
      Left            =   480
      TabIndex        =   0
      Top             =   1395
      Width           =   5130
      Begin VB.ComboBox cboMes 
         Height          =   315
         Left            =   90
         TabIndex        =   1
         Top             =   255
         Width           =   4950
      End
   End
End
Attribute VB_Name = "FrmLibroDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NombreTabla As String

Private Sub Form_Load()
  Call ConfiguraForm
  Call Llenacbomes
  NombreTabla = "CT_CABCOMPROB" & VGParamSistem.Anoproceso
  If IsNumeric(VGParamSistem.Anoproceso) Then
      Call SeleccionarMes(CInt(VGParamSistem.Mesproceso), CInt(VGParamSistem.Anoproceso))
  End If
End Sub

Sub ConfiguraForm()
  optOpcion(0).Value = True
  Ctr_Ayuda1(0).conexion VGCNx
  Ctr_Ayuda1(1).conexion VGCNx
  Ctr_Ayuda1(1).Filtro = "asientocodigo='" & Ctr_Ayuda1(0).xclave & "'"
  Ctr_Ayuda2(0).conexion VGCNx
  Ctr_Ayuda2(1).conexion VGCNx
  
End Sub

Private Sub optOpcion_Click(Index As Integer)
 Select Case Index
   Case 0:
     fraDetallado.Visible = True
     fraResumido.Visible = False
   
   Case 1:
     fraDetallado.Visible = False
     fraResumido.Visible = True
   
   Case 2:
     fraDetallado.Visible = False
     fraResumido.Visible = True
 End Select

End Sub

Private Sub cmdBotones_Click(Index As Integer)
 Dim cMes As String
  Select Case Index
    Case 0:
     'If ValidaData() = True Then
     '   MsgBox "Imprimiendo", vbInformation, Caption
     'End If
       Dim arrform(2) As Variant, arrparm() As Variant
       If optOpcion(0).Value = True Then   'Detallado
            ReDim arrparm(8)               'Store Procedure:CT_DIARIO2_REP
            arrparm(0) = VGParamSistem.BDEmpresa
            arrparm(1) = VGParametros.empresacodigo
            arrparm(2) = VGParamSistem.Anoproceso
            If cboMes.ListIndex >= 0 Then
                cMes = Format(cboMes.ListIndex + 1, "0#")
            Else
                cMes = Format(VGParamSistem.Mesproceso, "0#")
            End If
            arrparm(3) = cMes
            arrparm(4) = "%%"
            arrparm(5) = IIf(Ctr_Ayuda2(0).xclave = Empty, "%%", Ctr_Ayuda2(0).xclave)
            arrparm(6) = IIf(Ctr_Ayuda2(1).xclave = Empty, "%%", Ctr_Ayuda2(1).xclave)
            arrparm(7) = "FORMATO 05.01"
            Set VGvardllgen = New dllgeneral.dll_general
            arrform(0) = "@TituloReporte='" & "Libro Diario Detallado - Asiento: " & Ctr_Ayuda1(0).xclave & " " & Ctr_Ayuda1(0).xnombre & "'"
            arrform(1) = "@Mes='" & VGvardllgen.DesMes(cMes) & "'"     'VGvardllgen.DESMES(VGParamSistem.Mesproceso)
            Call ImpresionRptProc("ct_LibroDiario.rpt", arrform, arrparm)
       Else     'Resumido Store Procedure:CT_DIARIO1_REP
            ReDim arrparm(6)
            arrparm(0) = VGParamSistem.BDEmpresa
            arrparm(1) = VGParametros.empresacodigo
            arrparm(2) = VGParamSistem.Anoproceso
            If cboMes.ListIndex >= 0 Then
                cMes = Format(cboMes.ListIndex + 1, "0#")
            Else
                cMes = Format(VGParamSistem.Mesproceso, "0#")
            End If
            arrparm(3) = cMes
            arrparm(4) = IIf(Ctr_Ayuda1(0).xclave = Empty, "%%", Ctr_Ayuda1(0).xclave)
            arrparm(5) = IIf(Ctr_Ayuda1(1).xclave = Empty, "%%", Ctr_Ayuda1(1).xclave)
            Set VGvardllgen = New dllgeneral.dll_general
            arrform(0) = "@TituloReporte='" & "Libro Diario Resumido - Asiento: " & Ctr_Ayuda1(0).xclave & " " & Ctr_Ayuda1(0).xnombre & "'"
            arrform(1) = "@Mes='" & VGvardllgen.DesMes(cMes) & "'"     'VGvardllgen.DESMES(VGParamSistem.Mesproceso)
            If optOpcion(1).Value = True Then
               Call ImpresionRptProc("ct_LibroDiarioResumido1.rpt", arrform, arrparm)
            Else
               Call ImpresionRptProc("ct_LibroDiarioResumido2.rpt", arrform, arrparm)
            End If
       End If
    
    Case 1: Unload Me
  
  End Select
End Sub

Sub Llenacbomes()
 Dim i As Integer
 Set VGvardllgen = New dllgeneral.dll_general
 cboMes.Clear
 For i = 1 To 12
   cboMes.AddItem VGvardllgen.DesMes(Str(i))
 Next
 Set VGvardllgen = Nothing

End Sub

Sub SeleccionarMes(nMes As Integer, nAnno As Integer)
  cboMes.Text = cboMes.List(nMes - 1)
  DTPickerFecInicio.Value = Format("01/" & nMes & "/" & nAnno, "dd/mm/yyyy")
  DTPickerFecFinal.Value = DateAdd("d", -1, DateAdd("m", 1, DTPickerFecInicio.Value))
End Sub

Private Sub cboMes_Click()
  Call SeleccionarMes(cboMes.ListIndex + 1, CInt(VGParamSistem.Anoproceso))
End Sub

Function ValidaData() As Boolean
 Dim SQL As String
    Set VGvardllgen = New dllgeneral.dll_general
    If DTPickerFecInicio.Value > DTPickerFecFinal.Value Then
        MsgBox "La Fecha de Término es menor a la Fecha de Inicio", vbInformation, Caption
        DTPickerFecInicio.SetFocus
        ValidaData = False
        Exit Function
    End If
    
    SQL = "select name from sysobjects where name='" & NombreTabla & "'"
    If VGvardllgen.VerificaDatoExistente(VGCNx, SQL) > 0 Then
        SQL = "select asientocodigo from " & NombreTabla & " "
        If optOpcion(0).Value = True Then
            SQL = SQL & "WHERE  asientocodigo='" & Ctr_Ayuda2(0).xclave & "' "
        Else
            SQL = SQL & "WHERE asientocodigo='" & Ctr_Ayuda1(0).xclave & "' "
        End If
        SQL = SQL & " AND cabcomprobmes=" & CInt(VGParamSistem.Mesproceso)
        If VGvardllgen.VerificaDatoExistente(VGCNx, SQL) = 0 Then
            MsgBox "No existe Información para Mostrar", vbInformation, Caption
            ValidaData = False
            Exit Function
        End If
    Else
        MsgBox "No existen las Tablas para el Año de Proceso Actual: " & VGParamSistem.Anoproceso & vbCrLf & "Consulte con el Administrador del Sistema Contable", vbExclamation, Caption
        ValidaData = False
        Exit Function
    End If
          
    ValidaData = True
End Function

Private Sub DTPickerFecInicio_GotFocus()
    DTPickerFecInicio.CalendarBackColor = &HE2FDFE
End Sub

Private Sub DTPickerFecFinal_GotFocus()
    DTPickerFecFinal.CalendarBackColor = &HE2FDFE
End Sub

Private Sub Ctr_Ayuda1_AlDevolverDato(Index As Integer, ByVal ColecCampos As ADODB.Fields)
 If Index = 0 Then
   Ctr_Ayuda1(1).Filtro = "asientocodigo='" & Trim(Ctr_Ayuda1(0).xclave) & "'"
 End If

End Sub

Private Sub Ctr_Ayuda2_AlDevolverDato(Index As Integer, ByVal ColecCampos As ADODB.Fields)
 If Index = 0 Then
   Ctr_Ayuda2(1).Filtro = "asientocodigo='" & Trim(Ctr_Ayuda2(0).xclave) & "'"
 End If

End Sub


