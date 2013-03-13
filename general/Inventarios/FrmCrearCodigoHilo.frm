VERSION 5.00
Begin VB.Form FrmCrearCodigoHilo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crear Codigo de Hilo"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7440
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   14
      Top             =   0
      Width           =   7215
      Begin VB.ComboBox cbo_TipoHilo 
         Height          =   315
         Left            =   2520
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.ComboBox cbo_color 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cbo_cabos 
         Height          =   315
         Left            =   3240
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cbo_Capilares 
         Height          =   315
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cbo_Mezcla 
         Height          =   315
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cbo_Titulo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1200
         Width           =   1215
      End
      Begin VB.ComboBox cbo_Familia 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label9 
         Caption         =   "Tipo de Hilo :"
         Height          =   255
         Left            =   2520
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Color de Hilo :"
         Height          =   255
         Left            =   4920
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label6 
         Caption         =   "N° de Cabos :"
         Height          =   255
         Left            =   3240
         TabIndex        =   19
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "N° de Capilares :"
         Height          =   255
         Left            =   1560
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Mezcla:"
         Height          =   255
         Left            =   4920
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Título:"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Familia:"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   1680
      Width           =   7215
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
         TabIndex        =   13
         Top             =   720
         Width           =   6735
         WordWrap        =   -1  'True
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
         TabIndex        =   12
         Top             =   240
         Width           =   2295
      End
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
         TabIndex        =   11
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   7215
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
         TabIndex        =   9
         Top             =   240
         Width           =   1095
      End
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
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCrearCodigoHilo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Coneccion As New ADODB.Connection
Private RsHilo As New ADODB.Recordset

Dim SQL As String
Dim sqlEmpresa As String
Private Sub cbo_Ancho_Click()
    CreaCodigo
End Sub
Private Sub cbo_Densidad_Click()
    CreaCodigo
End Sub



Private Sub cbo_cabos_Click()
    CreaCodigo
End Sub

Private Sub cbo_Capilares_Change()
    CreaCodigo
End Sub



Private Sub cbo_Capilares_Click()
    CreaCodigo
End Sub

Private Sub cbo_color_Click()
    CreaCodigo
End Sub

Private Sub cbo_Familia_Click()
    CreaCodigo
End Sub
Private Sub cbo_Mezcla_Click()
    CreaCodigo
End Sub

Private Sub cbo_TipoHilo_Click()
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
    
    If Trim(cbo_TipoHilo.text) = "" Then
        ccadena = ccadena + "Tipo "
    End If
    
    If Trim(cbo_Mezcla.text) = "" Then
        ccadena = ccadena + "Mezcla "
    End If
    
    If Trim(cbo_Titulo.text) = "" Then
        ccadena = ccadena + "Título "
    End If
    
    If Trim(cbo_Capilares.text) = "" Then
        ccadena = ccadena + "Capilar "
    End If
    
    If Trim(cbo_cabos.text) = "" Then
        ccadena = ccadena + "Cabos "
    End If
    If Trim(cbo_color.text) = "" Then
        ccadena = ccadena + "Color "
    End If
    
    
    If Trim(ccadena) <> "" Or Trim(Label1.Caption) = "" Then If MsgBox("Debe Seleccionar: " + ccadena, vbCritical + vbOKOnly, "Error") = vbOK Then cbo_Familia.SetFocus: Exit Sub
    'SQL = "select telacrudaid from [Maestro Tela Cruda] where telacrudaid='" + Trim(Label1.Caption) + "'"
    SQL = "select acodigo from maeart where acodigo='" + Trim(Label1.Caption) + "'"
    
    RsHilo.Open SQL, VGcnx, adOpenStatic
    'Set RsHilo = Coneccion.Execute(SQL)
    
    Dim cuantos As Integer
    cuantos = 0
    While Not RsHilo.EOF
        cuantos = cuantos + 1
        RsHilo.MoveNext
    Wend
    If cuantos >= 1 Then
        MsgBox "El Código seleccionado ya existe en la Base de Datos.", vbInformation
        RsHilo.Close
        cbo_Familia.SetFocus
        Exit Sub
    End If
    
    RsHilo.Close
    
    Dim Fecha As Date
    'fecha = Date + Time()
    'SQL = "insert [Maestro Tela Cruda] (telacrudaid,telacrudadescripcion) values('" + Trim(Label1.Caption) + "','" + Trim(Label8.Caption) + "')"
    SQL = "INSERT INTO MAEART (acodigo,adescri,afamilia,amodelo,aunidad,agrupo,aflote,afserie) values('" + Trim(Label1.Caption) + "','" + Trim(Label8.Caption) + "', 'HI', '0', 'KG', '0', 'N', 'N')"
    
    'Coneccion.Execute (SQL)
    VGcnx.Execute (SQL)
    Label1.Caption = ""
    Label8.Caption = ""
    cbo_Familia.SetFocus
       
    MsgBox "Datos Grabados Satisfactoriamente ...", vbInformation
    CreaCodigo
    
End Sub
Private Sub cmd_Salir_Click()
    Set RsHilo = Nothing
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
'***********************
'   carga familia
    Set RsHilo = New ADODB.Recordset
    
    SQL = "select * from familia_Hilo order by codigo"
    RsHilo.Open SQL, VGcnx, adOpenStatic
    'Set RsHilo = Coneccion.Execute(SQL)
    While Not RsHilo.EOF
        cbo_Familia.AddItem RsHilo!descripcion + Space(100) + RsHilo!codigo
        RsHilo.MoveNext
    Wend
    RsHilo.Close
    
'   carga tipo de hilo
    SQL = "select * from Tipos_Hilo order by codigo"
    RsHilo.Open SQL, VGcnx, adOpenStatic
    'Set RsHilo = Coneccion.Execute(SQL)
    While Not RsHilo.EOF
        If IsNull(RsHilo!descripcion) Then
            cbo_TipoHilo.AddItem " " + Space(100) + RsHilo!codigo
        Else
            cbo_TipoHilo.AddItem RsHilo!descripcion + Space(100) + RsHilo!codigo
        End If
        RsHilo.MoveNext
    Wend
    RsHilo.Close

'   carga mezcla
    SQL = "select * from mezcla order by codigo"
    RsHilo.Open SQL, VGcnx, adOpenStatic
    'Set RsHilo = Coneccion.Execute(SQL)
    While Not RsHilo.EOF
        If IsNull(RsHilo!descripcio) Then
            cbo_Mezcla.AddItem " " + Space(100) + RsHilo!codigo
        Else
            cbo_Mezcla.AddItem RsHilo!descripcio + Space(100) + RsHilo!codigo
        End If
        RsHilo.MoveNext
    Wend
    RsHilo.Close


'   carga titulo
    SQL = "select * from Titulos_Hilo order by codigo"
    RsHilo.Open SQL, VGcnx, adOpenStatic
    'Set RsHilo = Coneccion.Execute(SQL)
    While Not RsHilo.EOF
        cbo_Titulo.AddItem RsHilo!descripcion + Space(100) + RsHilo!codigo
        RsHilo.MoveNext
    Wend
    RsHilo.Close
    
    
'   carga Nro Capilares
    SQL = "select * from Capilares_hilo order by codigo"
    RsHilo.Open SQL, VGcnx, adOpenStatic
    'Set RsHilo = Coneccion.Execute(SQL)
    While Not RsHilo.EOF
       If IsNull(RsHilo!descripcion) Then
            cbo_Capilares.AddItem " " + Space(100) + RsHilo!codigo
        Else
            cbo_Capilares.AddItem RsHilo!descripcion + Space(100) + RsHilo!codigo
        End If
        RsHilo.MoveNext
    Wend
    RsHilo.Close
    
'   carga Cabos
    SQL = "select * from Cabos_Hilo order by codigo"
    RsHilo.Open SQL, VGcnx, adOpenStatic
    'Set RsHilo = Coneccion.Execute(SQL)
    While Not RsHilo.EOF
        If RsHilo!descripcion = "" Then
            cbo_cabos.AddItem " " + Space(100) + RsHilo!codigo
        Else
            cbo_cabos.AddItem RsHilo!descripcion + Space(100) + RsHilo!codigo
        End If
        RsHilo.MoveNext
    Wend
   RsHilo.Close
   
'   carga Color hilo
    SQL = "select * from Color_Hilo order by codigo"
    RsHilo.Open SQL, VGcnx, adOpenStatic
    'Set RsHilo = Coneccion.Execute(SQL)
    While Not RsHilo.EOF
        If RsHilo!descripcion = "" Then
            cbo_color.AddItem " " + Space(100) + RsHilo!codigo
        Else
            cbo_color.AddItem RsHilo!descripcion + Space(100) + RsHilo!codigo
        End If
        RsHilo.MoveNext
    Wend
   RsHilo.Close
   
End Sub

Private Sub CreaCodigo()
    Label1.Caption = Right(cbo_Familia.text, 2) + Right(cbo_TipoHilo.text, 1) + Right(cbo_Mezcla.text, 1) + Right(cbo_Titulo.text, 3) + Right(cbo_Capilares.text, 1) + Right(cbo_cabos.text, 1) + Right(cbo_color.text, 2)
    Label8.Caption = Trim(Left(UCase(cbo_Familia.text), 80)) + " " + Trim(Left(UCase(cbo_TipoHilo.text), 80)) + " " + Trim(Left(UCase(cbo_Mezcla.text), 50)) + " " + Trim(Left(UCase(cbo_Titulo.text), 20)) + " " + Trim(Left(UCase(cbo_Capilares.text), 10)) + " " + Trim(Left(UCase(cbo_cabos.text), 10)) + " " + Trim(Left(UCase(cbo_color.text), 50))
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Coneccion.Close
End Sub


