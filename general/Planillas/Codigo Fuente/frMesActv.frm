VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frMesActv 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Meses Activos"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "frMesActv.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5400
   Tag             =   "Panel de meses activos de planilla "
   Begin VB.Frame xFrame 
      Caption         =   "Seleccione"
      Height          =   1485
      Left            =   2805
      TabIndex        =   4
      Top             =   2550
      Visible         =   0   'False
      Width           =   2475
      Begin VB.CommandButton cmBoton 
         Caption         =   "&Cancelar"
         Height          =   420
         Index           =   1
         Left            =   1365
         TabIndex        =   7
         Top             =   885
         Width           =   840
      End
      Begin VB.CommandButton cmBoton 
         Caption         =   "&Aceptar"
         Height          =   420
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   885
         Width           =   840
      End
      Begin MSComCtl2.DTPicker xMes 
         Height          =   330
         Left            =   165
         TabIndex        =   5
         Top             =   360
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   582
         _Version        =   393216
         CustomFormat    =   "MMM/yyyy"
         Format          =   23592963
         CurrentDate     =   36634
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2340
      Top             =   1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frMesActv.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frMesActv.frx":0796
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frMesActv.frx":0AB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMeses 
      Height          =   780
      Left            =   3030
      TabIndex        =   3
      Top             =   1425
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1376
      ButtonWidth     =   1032
      ButtonHeight    =   1376
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cerrar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Abrir"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.ListBox xLista 
      Height          =   2595
      Left            =   60
      TabIndex        =   2
      Top             =   1455
      Width           =   2580
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Meses Activos"
      Height          =   195
      Left            =   60
      TabIndex        =   1
      Top             =   1200
      Width           =   1035
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frMesActv.frx":0DCE
      ForeColor       =   &H8000000E&
      Height          =   825
      Left            =   765
      TabIndex        =   0
      Top             =   135
      Width           =   4545
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "frMesActv.frx":0EC6
      Top             =   165
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000002&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000002&
      Height          =   1035
      Left            =   0
      Top             =   0
      Width           =   5415
   End
End
Attribute VB_Name = "frMesActv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RSMESES As New ADODB.Recordset
Dim REGACT As REGWIN
Dim FLAG01 As Boolean
Private Sub CMBOTON_Click(INDEX As Integer)
On Error GoTo ERRFLAG
    If INDEX = 0 Then
        'IF RSMESES.RECORDCOUNT > 0 THEN
            'SI DESEA CREAR UNO NUEVO
            NomTabla = Trim("BOL" & Format(xMes.Month, "00") & Format(xMes.Year, "0000"))
            FLAG01 = False
            Dim RSNEW As ADODB.Recordset
            Set RSNEW = New ADODB.Recordset
            RSNEW.Open NomTabla, DBSYSTEM, adOpenKeyset, adLockOptimistic
            If FLAG01 Then
                'SE PROCEDEN A CREAR LAS TABLAS DEL SISTEMA PARA SU EDICIÓN
                'ESTAS TABLAS SE CREAN MES POR MES
                'CULMINADO EL AÑO SE TRASLADAN A LA BASE DE DATOS ALMCNPLN.MDB
                DBSYSTEM.Execute "CREATE TABLE " & NomTabla & " (INUMBOL INT IDENTITY, NUMBOL INT, CODNOMBOL INT, CODTRAB VARCHAR(8), FECHA DATETIME, CODAFP VARCHAR(2), CCOSTO VARCHAR(50), TIPOPLAN INT, BASICO  Numeric(20,2) , SUMAAFP  Numeric(20,2) , SUMASALUD  Numeric(20,2) , SUMAIES  Numeric(20,2) , SUMARENTA  Numeric(20,2) , SUMASCTR  Numeric(20,2) , SUMACTS  Numeric(20,2) , SUMAGRAT  Numeric(20,2) , SUMAVAC  Numeric(20,2) , TOTING  Numeric(20,2) , TOTEGR  Numeric(20,2) , HORASTRAB  Numeric(20,2) , HORASEXTRAS  Numeric(20,2) , RENTA5TA  Numeric(20,2) )"
                NomTabla = "MOV" & Format(xMes.Month, "00") & Format(xMes.Year, "0000")
                DBSYSTEM.Execute "CREATE TABLE " & NomTabla & " (INUMBOL INT, CONCEPTO VARCHAR(10), MONTO  Numeric(20,2) , CODNOMBOL INT)"
            End If
            Set RSNEW = Nothing
            xMes.Day = 1
            RSMESES.MoveFirst
            RSMESES.FIND "MESACTIVO=#" & xMes.Value & "#"
            If RSMESES.EOF Then
                RSMESES.AddNew
                RSMESES!MESACTIVO = xMes.Value
                RSMESES!FECHA = Date
                RSMESES!NOMBRE = AMESES(Month(xMes.Value)) & " DE " & Year(xMes.Value)
                RSMESES!CERRADO = 0
                RSMESES.Update
                CARGAMESES
            Else
                MsgBox "El mes ya se encuentra activo", vbCritical
            End If
        'END IF
    End If
    xFrame.Visible = True
    xLista.Visible = True
    tbMeses.Visible = True
    xFrame.Visible = False
    Exit Sub
ERRFLAG:
    FLAG01 = True
    Resume Next
    Resume
End Sub

Private Sub Form_Activate()
    ActivarTools REGACT
End Sub

Private Sub Form_Load()
    RSMESES.Open "MESESACT", DBSYSTEM, adOpenKeyset, adLockOptimistic
    CARGAMESES
    With REGACT
        .BUSCAR = False
        .EDITAR = False
        .ELIMINAR = False
        .NUEVO = False
        .FILTRAR = False
        .IMPRIMIR = False
        .PRELIMINAR = False
    End With
    
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSMESES = Nothing
End Sub

Private Sub TBMESES_BUTTONCLICK(ByVal Button As MSComctlLib.Button)
    Select Case Button.INDEX
        Case 1 'NUEVO
            xMes.Value = Format(Date, "MMM/yyyy")
            xFrame.Visible = True
            xLista.Visible = False
            tbMeses.Visible = False
        Case 3 'CERRAR
            If xLista.ListCount = 0 Or xLista.ListIndex = -1 Then
                MsgBox "No existe elemento seleccionado", vbCritical
                Exit Sub
            End If
            If MsgBox("Seguro de cerrar el mes: " & xLista.Text, vbYesNo + vbQuestion) = vbYes Then
                RSMESES.MoveFirst
                RSMESES.FIND "NOMBRE='" & xLista.Text & "'"
                RSMESES.Delete
                CARGAMESES
            End If
        Case 5 'ABRIR
            xFrame.Visible = True
            xLista.Visible = False
            tbMeses.Visible = False
    End Select
End Sub

Public Sub CARGAMESES()
    xLista.Clear
    If RSMESES.RecordCount = 0 Then Exit Sub
    RSMESES.MoveFirst
    Do While Not RSMESES.EOF
        xLista.AddItem RSMESES!NOMBRE
        RSMESES.MoveNext
    Loop
End Sub

