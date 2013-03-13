VERSION 5.00
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form frAddColTrab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adición de Datos Informativos"
   ClientHeight    =   3900
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frAddColTrab.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3900
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   2531
      TabIndex        =   8
      Top             =   3330
      Width           =   1290
   End
   Begin VB.CommandButton cmAgregar 
      Caption         =   "&Agregar"
      Height          =   360
      Left            =   859
      TabIndex        =   7
      Top             =   3330
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Caption         =   "Otro Dato Informativo (Adición)"
      Height          =   3045
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin AplisetControlText.Aplitext xDescripcion 
         Height          =   285
         Left            =   1755
         TabIndex        =   2
         Top             =   945
         Width           =   2475
         _ExtentX        =   4366
         _ExtentY        =   503
         MaxLength       =   30
         Text            =   ""
      End
      Begin AplisetControlText.Aplitext xCodigo 
         Height          =   285
         Left            =   1755
         TabIndex        =   1
         Top             =   435
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   503
         MaxLength       =   15
         Text            =   ""
         TipoCodigo      =   -1  'True
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Tipo Si o No"
         Height          =   225
         Left            =   1755
         TabIndex        =   6
         Top             =   2505
         Width           =   2235
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Tipo Fecha"
         Height          =   225
         Left            =   1755
         TabIndex        =   5
         Top             =   2235
         Width           =   2235
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Tipo Numérico"
         Height          =   225
         Left            =   1755
         TabIndex        =   4
         Top             =   1965
         Width           =   2235
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Tipo Alfanumérico (Texto)"
         Height          =   225
         Left            =   1755
         TabIndex        =   3
         Top             =   1695
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Información"
         Height          =   195
         Left            =   300
         TabIndex        =   11
         Top             =   1485
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Left            =   300
         TabIndex        =   10
         Top             =   975
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Left            =   300
         TabIndex        =   9
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frAddColTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub CMAGREGAR_CLICK()
    Dim X As Long
    If xCodigo.Text = "" Then
        MsgBox "Falta ingresar un codigo", vbInformation
        xCodigo.SetFocus
        Exit Sub
    End If
    If xDescripcion.Text = "" Then
        MsgBox "Falta ingresar una descripción", vbInformation
        xDescripcion.SetFocus
        Exit Sub
    End If
    DBSYSTEM.Execute "UPDATE DATATRAB SET CODDATA=CODDATA WHERE CODDATA='" & xCodigo.Text & "'", X
    If X <> 0 Then
        MsgBox "El codigo ingresado ya existe, por favor cambielo o cancele la operación. La operación no ha sido completada", vbInformation
        Exit Sub
    End If
    DBSYSTEM.Execute "UPDATE DATATRAB SET CODDATA=CODDATA WHERE DESCDATA='" & xDescripcion.Text & "'", X
    If X <> 0 Then
        MsgBox "La descripción ingresada ya existe, por favor cambielo o cancele la operación. La operación no ha sido completada", vbInformation
        Exit Sub
    End If
    'VALIDANDO SI EL DATO SE ENCUENTRA YA EN CONCEPTOS DE REMUNERACIONES
    Dim RSVAL As New ADODB.Recordset
    RSVAL.Open "SELECT * FROM CONCEPTOS WHERE CODIGO='" & Trim(xCodigo.Text) & "'", DBSYSTEM, adOpenKeyset, adLockReadOnly
    If RSVAL.RecordCount > 0 Then
        MsgBox "Conflicto de NOMBRE este codigo de concepto ya se encuentra ingresado en concepto de remuneraciones " & Chr(13) & _
        "los codigos tienen que ser diferentes a los que existen en conceptos de remuneraciones", vbExclamation
        Exit Sub
    End If
    'VALIDANDO SI ES QUE SE ENCUENTRA EN LA TABLA VARIABLES
    Set RSVAL = Nothing
    RSVAL.Open "SELECT * FROM VARIABLES2 WHERE CODIGO='" & Trim(xCodigo.Text) & "'", DBSTARPLAN, adOpenKeyset, adLockReadOnly
    If RSVAL.RecordCount > 0 Then
        MsgBox "Usted esta usando un NOMBRE que corresponde a las variables del sistema", vbExclamation
        Exit Sub
    End If
    Dim xTipo As String
    If Option1.Value Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD " & xCodigo.Text & " VARCHAR(30) NULL"
        xTipo = "T"
    End If
    If Option2.Value Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD " & xCodigo.Text & "  Numeric(20,2)  NULL DEFAULT 0"
        xTipo = "N"
    End If
    If Option3.Value Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD " & xCodigo.Text & " DATETIME NULL"
        xTipo = "F"
    End If
    If Option4.Value Then
        DBSYSTEM.Execute "ALTER TABLE TRABAJADORES ADD " & xCodigo.Text & " BIT NULL DEFAULT 0"
        xTipo = "B"
    End If
    DBSYSTEM.Execute "INSERT INTO DATATRAB (CODDATA,DESCDATA,TIPODATA) VALUES ('" & xCodigo.Text & "','" & xDescripcion.Text & "','" & xTipo & "')"
    Unload Me
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub


Private Sub xCodigo_LostFocus()
    Dim CADENA As String
    CADENA = "'CODTRAB','NOMBRES','CODAREA','CODCCOSTO','BASICO','ASIGFAM','CODAFP','TASASCTR','APOROBL','SEGURO','TOPESEGURO'," & _
             "'COMISIONRA','SUMAAFP','SUMASALUD','TOTING','TOTEGR','_HORAST','_HOREXTRAS','_QUINTACAT','SUMAIES','SUMARENTA'," & _
             "'SUMASCTR','SUMACTS','SUMAGRAT','SUMAVAC','T1','T2','T3','T4','T5','OTROSING','OTROSEGR','ADELANTO','UBIGEO'," & _
             "'SEXO','TIPOTRAB','FECHAING','SITUACION','CARGO','BANCO','ESSALUDVIDA','RUCEPS','NOPDT','OPCION01','OPCION02'," & _
             "'OPCIONA','OPCIONB','XREDONDEO','AFECTOQUINTA'"
    If InStr(CADENA, "'" & Trim(xCodigo.Text) & "'") > 0 Then
        MsgBox "El codigo : " & xCodigo.Text & " es palabra reservada del sistema ", vbExclamation
        xCodigo.SetFocus
        Exit Sub
    End If
End Sub
