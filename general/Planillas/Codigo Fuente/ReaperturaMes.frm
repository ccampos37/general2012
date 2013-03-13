VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form ReaperturaMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reapertura de Mes"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "ReaperturaMes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "C&ancelar"
      Height          =   375
      Left            =   3450
      TabIndex        =   9
      Top             =   3435
      Width           =   1425
   End
   Begin VB.CommandButton cmContinuar 
      Caption         =   "&Reaperturar"
      Height          =   375
      Left            =   1815
      TabIndex        =   8
      Top             =   3435
      Width           =   1425
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ReaperturaMes.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Especificaciones de Ingreso"
      Height          =   2385
      Left            =   105
      TabIndex        =   0
      Top             =   885
      Width           =   6195
      Begin MSComCtl2.DTPicker xFechaFin 
         Height          =   285
         Left            =   1125
         TabIndex        =   7
         Top             =   1920
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24444929
         CurrentDate     =   36699
      End
      Begin MSComCtl2.DTPicker xFechaIni 
         Height          =   285
         Left            =   1125
         TabIndex        =   5
         Top             =   1575
         Visible         =   0   'False
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   24444929
         CurrentDate     =   36699
      End
      Begin MSComctlLib.ListView Lista 
         Height          =   1725
         Left            =   2535
         TabIndex        =   3
         Top             =   495
         Width           =   3540
         _ExtentX        =   6244
         _ExtentY        =   3043
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Periodos en Cronogramas"
            Object.Width           =   5733
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "FechaIni"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "FechaFin"
            Object.Width           =   2540
         EndProperty
      End
      Begin AplisetControlText.Aplitext xMes 
         Height          =   285
         Left            =   165
         TabIndex        =   2
         Top             =   510
         Width           =   2280
         _ExtentX        =   4022
         _ExtentY        =   503
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Periodos en Cronograma"
         Height          =   195
         Left            =   2580
         TabIndex        =   10
         Top             =   270
         Width           =   1740
      End
      Begin VB.Label l2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   165
         TabIndex        =   6
         Top             =   1980
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label l1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Inicial"
         Height          =   195
         Left            =   165
         TabIndex        =   4
         Top             =   1620
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Mes de Trabajo"
         Height          =   195
         Left            =   165
         TabIndex        =   1
         Top             =   270
         Width           =   1110
      End
   End
End
Attribute VB_Name = "ReaperturaMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSAUX As New ADODB.Recordset
Dim XITEM As ListItem
Dim XMESSTRING As String
Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub CMCONTINUAR_CLICK()
Dim X As Integer
If Len(xMes.Tag) > 0 Then
    If Len(XMESSTRING) > 0 Then
        If MsgBox("REALMENTE DESEA REAPERATURAR EL " & XMESSTRING & " .PLANILLA SELECCIONADA EN LA LISTA ?", vbYesNo, "CONFIRMACIÓN") = vbYes Then
                DBSYSTEM.Execute "UPDATE NOMBOL SET CERRADO=0 WHERE NOMBRE='" & XMESSTRING & "' AND FECHAINI=" & DateSQL(xFechaIni.Value) & " AND FECHAFIN=" & DateSQL(xFechaFin.Value) & "", X
                If X > 0 Then
                    MsgBox "LA PLANILLA " & XMESSTRING & " FUE REAPERTURADA SATISFACTORIAMENTE", vbInformation, "CLAUSURA DE MES"
                End If
        End If
    Else
        MsgBox "SELECCIONE LA PLANILLA A REAPERTURAR EN LA LISTA DEL MES DE " & Format(xMes.Tag, "MMMM"), vbInformation
        Exit Sub
    End If
Else
    MsgBox "SELECCIONE EL MES A REAPERTURAR", vbInformation
    xMes.SetFocus
    Exit Sub
End If
End Sub

Private Sub FORM_UNLOAD(Cancel As Integer)
    Set RSAUX = Nothing
End Sub
Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
    xFechaIni.Visible = True
    xFechaFin.Visible = True
    l1.Visible = True
    l2.Visible = True
    xFechaIni.Value = CDate(Item.SubItems(1))
    xFechaFin.Value = CDate(Item.SubItems(2))
    XMESSTRING = Item.Text
End Sub

Private Sub XMES_DBLCLICK()
    Lista.ListItems.Clear
    Dim RSMESES As New ADODB.Recordset
    Set RSMESES = New ADODB.Recordset
    RSMESES.Open "SELECT MESACTIVO, NOMBRE FROM MESESACT ORDER BY MESACTIVO", DBSYSTEM, adOpenStatic
    If RSMESES.RecordCount = 0 Then
        MsgBox "NO SE HAN ENCONTRADO MESES DESACTIVOS", vbCritical
        Set RSMESES = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSMESES
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xMes.Text = RSMESES!NOMBRE
        xMes.Tag = RSMESES!MESACTIVO
    Else
        Set RSMESES = Nothing
        Exit Sub
    End If
    Set RSMESES = Nothing
    'RECICLAJE DE RSMESES
    RSMESES.Open "SELECT CODIGO, NOMBRE, FECHAINI, FECHAFIN FROM NOMBOL WHERE CERRADO<>0 AND MES=" & DateSQL(CDate(xMes.Tag)) & " ORDER BY FECHAINI", DBSYSTEM, adOpenStatic
    Do While Not RSMESES.EOF
        Set XITEM = Lista.ListItems.Add(, , RSMESES!NOMBRE, , 1)
        XITEM.SubItems(1) = RSMESES!FECHAINI
        XITEM.SubItems(2) = RSMESES!FECHAFIN
        XITEM.Tag = RSMESES!CODIGO
        RSMESES.MoveNext
    Loop
    XMESSTRING = ""
    l1.Visible = False
    l2.Visible = False
    xFechaIni.Visible = False
    xFechaFin.Visible = False
    Set RSMESES = Nothing
End Sub

