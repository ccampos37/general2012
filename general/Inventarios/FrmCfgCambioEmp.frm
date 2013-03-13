VERSION 5.00
Begin VB.Form FrmCfgCambioEmp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selección de Empresa"
   ClientHeight    =   2685
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4965
   Icon            =   "FrmCfgCambioEmp.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2685
   ScaleWidth      =   4965
   Begin VB.Frame Frame1 
      Height          =   810
      Left            =   135
      TabIndex        =   2
      Top             =   1020
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1275
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   315
         Width           =   3225
      End
      Begin VB.Label Label3 
         Caption         =   "Empresa     :"
         Height          =   255
         Left            =   210
         TabIndex        =   4
         Top             =   330
         Width           =   975
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   3120
      Picture         =   "FrmCfgCambioEmp.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1920
      Width           =   775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   675
      Left            =   1560
      Picture         =   "FrmCfgCambioEmp.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1920
      Width           =   775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FrmCfgCambioEmp.frx":114E
      Height          =   825
      Left            =   120
      TabIndex        =   5
      Top             =   150
      Width           =   4695
   End
End
Attribute VB_Name = "FrmCfgCambioEmp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cR As ADODB.Recordset, cS As String
Dim adoreg As ADODB.Recordset
Private Sub Command1_Click() 'ACEPTAR
Dim RSQL As String
Dim rs As ADODB.Recordset
Dim AdUs As ADODB.Recordset
Dim IASA As String
If cR.RecordCount > 0 Then
    
    cR.MoveFirst
    cR.Move (Combo1.ListIndex)
    'vGNomRep = cR.Fields("EMP_REPORTE")
    'MDIMenu.StatusBar1.Panels(1).text = vGNomEmp
    If vGAdmLog Then
        VGCODEMPRESA = cR.Fields(0)
        'VERIFICA QUE SE ENCUENTRA LA BASE DE DATOS
        If UCase(Dir$(sName & "Data\" & VGCODEMPRESA & "\" & "BdComun.mdb")) <> "BDCOMUN.MDB" Then
              If MsgBox("No se encuentra la base de datos necesaria para esta empresa ;" & Chr(13) & "Desea Crearla ", vbQuestion + vbOKCancel) = vbOK Then
                 If Not AGREGARBASE(VGCODEMPRESA) Then Screen.MousePointer = 1: Exit Sub
               Else:
                  Exit Sub
              End If
        End If
        

        HabilitarMenu_Usuarios VGUsuario, VGCODEMPRESA, "A2"
    Else
        Set AdUs = New ADODB.Recordset
        AdUs.Open "Select * From Usuario_Inv Where Usu_Codigo = '" & VGUsuario & "' and  Emp_Codigo = '" & cR.Fields(0) & "' ", VGconfig, adOpenStatic
        If AdUs.RecordCount > 0 Then
                If UCase(RTrim$(VGPass)) = DECODIFICA(AdUs.Fields("USU_PASSWORD"), NUMMAGICO) Then
                            VGCODEMPRESA = cR.Fields(0)
                            HabilitarMenu_Usuarios VGUsuario, VGCODEMPRESA
                End If
               
        Else
                MsgBox "El Usuario no tiene Acceso a la Empresa   " & cR.Fields("EMP_RAZON_NOMBRE"), vbInformation, "Información"
        End If
        AdUs.Close
    End If
    If UCase(Dir$(cRuta4)) = UCase(VGNameCont & ".MDB") Then
     'VGcnxCT.Close
    End If
    cRuta4 = VGParamSistem.RutaReport & VGNameCont & ".MDB"                         'cNomBd4      'BD. Contabilidad
    cNomBd2 = "BDComun.mdb"
    VGCNx.CursorLocation = adUseClient
    'Vgcnx.ConnectionString = "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=" & cRuta2 & ";"
    
  '  VGCNx.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=sa;password='" & Trim(VGPassw) & "';Initial Catalog=" & VGParamSistem.BDEmpresa & ";Data Source=" & VGServer
    
  '  VGCNx.Open
    'Carga la configuracion
    RSQL = "select * from configuracion"
    Set adoreg = New ADODB.Recordset
    adoreg.Open RSQL, VGCNx, adOpenDynamic, adLockOptimistic
    MDIPrincipal.mnu_ajuste.Visible = False
    MDIPrincipal.mnu_artven_06.Visible = False
    VGAutomatico = False
    If Not adoreg.EOF Then
        IASA = IIf(IsNull(adoreg("cod_iasa")), "", adoreg("cod_iasa"))
        If IASA <> "" Then
                       'MDIPrincipal.mnu_repIASA.Visible = True
                       'MDIPrincipal.mnu_guiaIngIasa.Visible = True
                       'MDIPrincipal.mnu_recepcion.Visible = True
                       VGIASA = IASA
       End If
       VGAutomatico = IIf(adoreg("cod_bloqueo"), True, False)
       If adoreg("cod_bloqueo") Then
'          MDIPrincipal.mnu_Asiento_02.Visible = True
       End If
       
  End If
  adoreg.Close
 'el almacen por defecto
    RSQL = "Select  * From  TabAlm"
    Set rs = New ADODB.Recordset
    rs.Open RSQL, VGCNx, adOpenStatic
    If rs.RecordCount = 0 Then
        VGNomAlm = " "
    Else
        VGNomAlm = IIf(IsNull(rs("tadescri")), " ", rs("tadescri"))
        VGAlma = rs("taalma")
    End If
    rs.Close
    MDIPrincipal.Caption = "Sistema de Inventario" & "     " & VGNomAlm & "    " & VGparametros.RucEmpresa
    'Set VGWrk = Workspaces(0)          'Cambiar todo Ado
    'Set VGBaseDatos = VGWrk.OpenDatabase(cRuta2, False, False)
    'Set VGBaseDatos = VGWrk.OpenDatabase(cRuta2, False, False)
    
End If
cR.Close
Command2_Click
End Sub
Private Sub Command2_Click() 'SALIR
Unload Me
End Sub

Private Sub Form_Activate()
 Combo1.SetFocus
End Sub

Private Sub Form_Load()
central Me
IniObj

End Sub

Private Sub IniObj()
Combo1.Clear
cS = "Select EMP_CODIGO,EMP_RAZON_NOMBRE,EMP_REPORTE  From EMPRESA  order by EMP_CODIGO"
Set cR = New ADODB.Recordset
cR.Open cS, VGconfig, adOpenStatic
Do While Not cR.EOF
    If Not IsNull(cR.Fields("EMP_RAZON_NOMBRE")) Then
        Combo1.AddItem cR.Fields("EMP_RAZON_NOMBRE")
    End If
    cR.MoveNext
    If cR.EOF Then Exit Do
Loop
If Combo1.ListCount <> 0 Then
     Combo1.ListIndex = 0
End If
End Sub
