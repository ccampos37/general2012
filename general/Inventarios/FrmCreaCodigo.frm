VERSION 5.00
Begin VB.Form FrmCreaCodigo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Creación de Código de Tela"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   17
      Top             =   3360
      Width           =   7215
      Begin VB.CommandButton cmd_Grabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   6
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmd_Salir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4320
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   7215
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   2520
         TabIndex        =   16
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label7 
         Caption         =   "Código Generado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   615
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   6735
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox cbo_Familia 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cbo_Titulo 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cbo_Mezcla 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cbo_Ancho 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cbo_Densidad 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   2055
      End
      Begin VB.Label Label2 
         Caption         =   "Familia:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Título:"
         Height          =   255
         Left            =   2520
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Mezcla:"
         Height          =   255
         Left            =   4920
         TabIndex        =   10
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Ancho:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Densidad:"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmCreaCodigo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Coneccion As New ADODB.Connection
Private RsTela As New ADODB.Recordset

Dim SQL As String
Dim sqlEmpresa As String
Private Sub cbo_Ancho_Click()
    CreaCodigo
End Sub
Private Sub cbo_Densidad_Click()
    CreaCodigo
End Sub
Private Sub cbo_Familia_Click()
    CreaCodigo
End Sub
Private Sub cbo_Mezcla_Click()
    CreaCodigo
End Sub
Private Sub cbo_Titulo_Click()
    CreaCodigo
End Sub
Private Sub cmd_Grabar_Click()
    Dim ccadena As String
    ccadena = ""
    If Trim(cbo_Familia.text) = "" Then
        ccadena = "Familia "
    End If
    If Trim(cbo_Titulo.text) = "" Then
        ccadena = ccadena + "Título "
    End If
    If Trim(cbo_Mezcla.text) = "" Then
        ccadena = ccadena + "Mezcla "
    End If
    If Trim(cbo_Ancho.text) = "" Then
        ccadena = ccadena + "Ancho "
    End If
    If Trim(cbo_Densidad.text) = "" Then
        ccadena = ccadena + "Densidad"
    End If
    If Trim(ccadena) <> "" Or Trim(Label1.Caption) = "" Then If MsgBox("Debe Seleccionar: " + ccadena, vbCritical + vbOKOnly, "Error") = vbOK Then cbo_Familia.SetFocus: Exit Sub
    'SQL = "select telacrudaid from [Maestro Tela Cruda] where telacrudaid='" + Trim(Label1.Caption) + "'"
    SQL = "select acodigo from maeart where acodigo='" + Trim(Label1.Caption) + "'"
    
    RsTela.Open SQL, VGCNx, adOpenStatic
    'Set RsTela = Coneccion.Execute(SQL)
    
    Dim cuantos As Integer
    cuantos = 0
    While Not RsTela.EOF
        cuantos = cuantos + 1
        RsTela.MoveNext
    Wend
    If cuantos >= 1 Then
        MsgBox "El Código seleccionado ya existe en la Base de Datos.", vbInformation
        RsTela.Close
        cbo_Familia.SetFocus
        Exit Sub
    End If
    
    RsTela.Close
    
    Dim Fecha As Date
    'fecha = Date + Time()
    'SQL = "insert [Maestro Tela Cruda] (telacrudaid,telacrudadescripcion) values('" + Trim(Label1.Caption) + "','" + Trim(Label8.Caption) + "')"
    SQL = "INSERT INTO MAEART (acodigo,adescri,afamilia,amodelo,aunidad,agrupo,aflote,afserie) values('" + Trim(Label1.Caption) + "','" + Trim(Label8.Caption) + "', 'CR', '0', 'KG', '0', 'N', 'N')"
    
    'Coneccion.Execute (SQL)
    VGCNx.Execute (SQL)
    Label1.Caption = ""
    Label8.Caption = ""
    cbo_Familia.SetFocus
       
    MsgBox "Datos Grabados Satisfactoriamente ...", vbInformation
    CreaCodigo
    
End Sub
Private Sub cmd_Salir_Click()
    Set RsTela = Nothing
    Unload Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
Tabula (KeyAscii)
End Sub

Private Sub Form_Load()
    
    
    central Me
''    CentrarFormulario Me
'    Open "c:\camtex.ini" For Input As #1
'        Input #1, xmServidor
'        Input #1, xmUser
'        Input #1, xmPass
'    Close #1

    'sqlEmpresa = "Provider=SQLOLEDB.1;Password=" & xmPass & ";Persist Security Info=true;User Id=" & xmUser & ";Initial Catalog=Empresas;Data Source=" & xmServidor
    'Coneccion.ConnectionString = sqlEmpresa
    'Coneccion.Open
    
'   carga familia
    Set RsTela = New ADODB.Recordset
    
    SQL = "select * from familiaTela order by codigo"
    RsTela.Open SQL, VGCNx, adOpenStatic
    'Set RsTela = Coneccion.Execute(SQL)
    While Not RsTela.EOF
        cbo_Familia.AddItem RsTela!descripcio + Space(100) + RsTela!codigo
        RsTela.MoveNext
    Wend
    RsTela.Close
'   carga titulo
    SQL = "select * from titulo order by codigo"
    RsTela.Open SQL, VGCNx, adOpenStatic
    'Set RsTela = Coneccion.Execute(SQL)
    While Not RsTela.EOF
        cbo_Titulo.AddItem RsTela!descripcio + Space(100) + RsTela!codigo
        RsTela.MoveNext
    Wend
    
    RsTela.Close
    
'   carga mezcla
    SQL = "select * from mezcla order by codigo"
    RsTela.Open SQL, VGCNx, adOpenStatic
    'Set RsTela = Coneccion.Execute(SQL)
    While Not RsTela.EOF
        If IsNull(RsTela!descripcio) Then
            cbo_Mezcla.AddItem " " + Space(100) + RsTela!codigo
        Else
            cbo_Mezcla.AddItem RsTela!descripcio + Space(100) + RsTela!codigo
        End If
        RsTela.MoveNext
    Wend
    RsTela.Close
    
'   carga ancho
    SQL = "select * from ancho_medida order by codigo"
    RsTela.Open SQL, VGCNx, adOpenStatic
    'Set RsTela = Coneccion.Execute(SQL)
    While Not RsTela.EOF
        cbo_Ancho.AddItem RsTela!descripcio + Space(100) + RsTela!codigo
        RsTela.MoveNext
    Wend
    RsTela.Close
    
'   carga densidad
    SQL = "select * from densidad_raport order by codigo"
    RsTela.Open SQL, VGCNx, adOpenStatic
    'Set RsTela = Coneccion.Execute(SQL)
    While Not RsTela.EOF
        If RsTela!descripcio = "" Then
            cbo_Densidad.AddItem " " + Space(100) + RsTela!codigo
        Else
            cbo_Densidad.AddItem RsTela!descripcio + Space(100) + RsTela!codigo
        End If
        RsTela.MoveNext
    Wend
   
   RsTela.Close
   
   
End Sub

Private Sub CreaCodigo()
    Label1.Caption = Right(cbo_Familia.text, 2) + Right(cbo_Titulo.text, 2) + Right(cbo_Mezcla.text, 1) + Right(cbo_Ancho.text, 3) + Right(cbo_Densidad.text, 2)
    Label8.Caption = Trim(Left(UCase(cbo_Familia.text), 80)) + " " + Trim(Left(UCase(cbo_Titulo.text), 80)) + " " + Trim(Left(UCase(cbo_Mezcla.text), 80)) + " " + Trim(Left(UCase(cbo_Ancho.text), 80)) + " " + Trim(Left(UCase(cbo_Densidad.text), 80))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Coneccion.Close
End Sub

