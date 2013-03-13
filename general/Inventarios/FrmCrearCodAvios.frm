VERSION 5.00
Begin VB.Form FrmCrearCodAvios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crearcion de Codigo de Avios"
   ClientHeight    =   4440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7395
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   120
      TabIndex        =   21
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
         TabIndex        =   15
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
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   18
      Top             =   1680
      Width           =   7215
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   735
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   840
         Width           =   6735
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   480
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   2895
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
         TabIndex        =   20
         Top             =   240
         Visible         =   0   'False
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
         TabIndex        =   12
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
         TabIndex        =   19
         Top             =   840
         Visible         =   0   'False
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
         TabIndex        =   2
         Top             =   480
         Width           =   2295
      End
      Begin VB.ComboBox cbo_Caracteristica 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cbo_Material 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   480
         Width           =   1815
      End
      Begin VB.ComboBox cbo_Medida 
         Height          =   315
         Left            =   2880
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cbo_color 
         Height          =   315
         Left            =   4920
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1200
         Width           =   2055
      End
      Begin VB.ComboBox cbo_Origen 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Familia:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Caracteristica :"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label4 
         Caption         =   "Material :"
         Height          =   255
         Left            =   5040
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Medida :"
         Height          =   255
         Left            =   3000
         TabIndex        =   8
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label10 
         Caption         =   "Color :"
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Origen :"
         Height          =   255
         Left            =   2880
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
   End
End
Attribute VB_Name = "FrmCrearCodAvios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Coneccion As New ADODB.Connection
Private RsAvios As New ADODB.Recordset

Dim SQL As String
Dim sqlEmpresa As String



Private Sub cbo_Medida_Change()
    CreaCodigo
End Sub

Private Sub cbo_Medida_Click()
    CreaCodigo
    
End Sub

Private Sub cbo_color_Click()
    CreaCodigo
End Sub

Private Sub cbo_Familia_Click()
    
    Cargar_Caracteristica
    Cargar_Medida
    CreaCodigo
End Sub
Private Sub cbo_Material_Click()
    CreaCodigo
End Sub

Private Sub cbo_Origen_Click()
    CreaCodigo
End Sub

Private Sub cbo_Caracteristica_Click()
    CreaCodigo
End Sub
Private Sub cmd_Grabar_Click()
    Dim ccadena As String
    ccadena = ""
    If Trim(cbo_Familia.text) = "" Then
        ccadena = "Familia "
    End If
    
    If Trim(cbo_Origen.text) = "" Then
        ccadena = ccadena + "Origen "
    End If
    
    If Trim(cbo_Material.text) = "" Then
        ccadena = ccadena + "Material "
    End If
    
    If Trim(cbo_Caracteristica.text) = "" Then
        ccadena = ccadena + "Caracteristica "
    End If
    
    If Trim(cbo_Medida.text) = "" Then
        ccadena = ccadena + "Medida "
    End If
    
    If Trim(cbo_color.text) = "" Then
        ccadena = ccadena + "Color "
    End If
    
    
    'If Trim(ccadena) <> "" Or Trim(Label1.Caption) = "" Then If MsgBox("Debe Seleccionar: " + ccadena, vbCritical + vbOKOnly, "Error") = vbOK Then cbo_Familia.SetFocus: Exit Sub
    If Trim(ccadena) <> "" Or Trim(Text1.text) = "" Then If MsgBox("Debe Seleccionar: " + ccadena, vbCritical + vbOKOnly, "Error") = vbOK Then cbo_Familia.SetFocus: Exit Sub
    'SQL = "select telacrudaid from [Maestro Tela Cruda] where telacrudaid='" + Trim(Label1.Caption) + "'"
    'SQL = "select acodigo from maeart where acodigo='" + Trim(Label1.Caption) + "'"
    SQL = "select acodigo from maeart where acodigo='" + Trim(Text1.text) + "'"
    
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    
    Dim cuantos As Integer
    cuantos = 0
    While Not RsAvios.EOF
        cuantos = cuantos + 1
        RsAvios.MoveNext
    Wend
    If cuantos >= 1 Then
        MsgBox "El Código seleccionado ya existe en la Base de Datos.", vbInformation
        RsAvios.Close
        cbo_Familia.SetFocus
        Exit Sub
    End If
    
    RsAvios.Close
    
    Dim Fecha As Date
    'fecha = Date + Time()
    'SQL = "insert [Maestro Tela Cruda] (telacrudaid,telacrudadescripcion) values('" + Trim(Label1.Caption) + "','" + Trim(Label8.Caption) + "')"
    'SQL = "INSERT INTO MAEART (acodigo,adescri,afamilia,amodelo,aunidad,agrupo,aflote,afserie) values('" + Trim(Label1.Caption) + "','" + Trim(Label8.Caption) + "', 'AV', '0', 'KG', '0', 'N', 'N')"
    SQL = "INSERT INTO MAEART (acodigo,adescri,afamilia,amodelo,aunidad,agrupo,aflote,afserie) values('" + Trim(Text1.text) + "','" + Trim(Text2.text) + "', 'AV', '0', 'KG', '0', 'N', 'N')"
    'Coneccion.Execute (SQL)
    VGcnx.Execute (SQL)
    Label1.Caption = ""
    Label8.Caption = ""
    cbo_Familia.SetFocus
       
    MsgBox "Datos Grabados Satisfactoriamente ...", vbInformation
    CreaCodigo
    
End Sub
Private Sub cmd_Salir_Click()
    Set RsAvios = Nothing
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
    Set RsAvios = New ADODB.Recordset
    
    SQL = "select * from familia_A order by descripcion"
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    While Not RsAvios.EOF
        cbo_Familia.AddItem RsAvios!descripcion + Space(100) + RsAvios!CodFam
        RsAvios.MoveNext
    Wend
    RsAvios.Close
    
'   carga origen
    SQL = "select * from Origen_A order by codori"
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    While Not RsAvios.EOF
        If IsNull(RsAvios!descripcion) Then
            cbo_Origen.AddItem " " + Space(100) + RsAvios!Codori
        Else
            cbo_Origen.AddItem RsAvios!descripcion + Space(100) + RsAvios!Codori
        End If
        RsAvios.MoveNext
    Wend
    RsAvios.Close

'   carga calidad
    SQL = "select * from calidad_a order by codmat"
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    While Not RsAvios.EOF
        If IsNull(RsAvios!descripcion) Then
            cbo_Material.AddItem " " + Space(100) + RsAvios!Codmat
        Else
            cbo_Material.AddItem RsAvios!descripcion + Space(100) + RsAvios!Codmat
        End If
        RsAvios.MoveNext
    Wend
    RsAvios.Close


'   carga caracteristica
    SQL = "select * from caracteristica_a order by codcar,codfam"
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    While Not RsAvios.EOF
    If IsNull(RsAvios!descripcion) Then
            cbo_Caracteristica.AddItem " " + Space(100) + RsAvios!Codcar
        Else
            cbo_Caracteristica.AddItem RsAvios!descripcion + Space(100) + RsAvios!Codcar
    End If
        RsAvios.MoveNext
    Wend
    RsAvios.Close
    
    
'   carga medida
    SQL = "select * from Medida_a order by codmed,codfam"
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    While Not RsAvios.EOF
       If IsNull(RsAvios!descripcion) Then
            cbo_Medida.AddItem " " + Space(100) + RsAvios!Codmed
        Else
            cbo_Medida.AddItem RsAvios!descripcion + Space(100) + RsAvios!Codmed
        End If
        RsAvios.MoveNext
    Wend
    RsAvios.Close
   
'   carga Color hilo
    SQL = "select * from Color_a order by codcol"
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    While Not RsAvios.EOF
        If RsAvios!descripcion = "" Then
            cbo_color.AddItem " " + Space(100) + RsAvios!Codcol
        Else
            cbo_color.AddItem RsAvios!descripcion + Space(100) + RsAvios!Codcol
        End If
        RsAvios.MoveNext
    Wend
   RsAvios.Close
   
End Sub

Private Sub CreaCodigo()
    'Label1.Caption = Right(cbo_Familia.text, 2) + Right(cbo_Origen.text, 1) + Right(cbo_Material.text, 1) + Right(cbo_Caracteristica.text, 2) + Right(cbo_Medida.text, 3) + Right(cbo_color.text, 3)
    'Label8.Caption = Trim(Left(UCase(cbo_Familia.text), 80)) + " " + Trim(Left(UCase(cbo_Origen.text), 80)) + " " + Trim(Left(UCase(cbo_Material.text), 50)) + " " + Trim(Left(UCase(cbo_Caracteristica.text), 80)) + " " + Trim(Left(UCase(cbo_Medida.text), 80)) + " " + Trim(Left(UCase(cbo_color.text), 80))
    
    Text1.text = Right(cbo_Familia.text, 2) + Right(cbo_Origen.text, 1) + Right(cbo_Material.text, 1) + Right(cbo_Caracteristica.text, 2) + Right(cbo_Medida.text, 3) + Right(cbo_color.text, 3)
    Text2.text = Trim(Left(UCase(cbo_Familia.text), 80)) + " " + Trim(Left(UCase(cbo_Origen.text), 80)) + " " + Trim(Left(UCase(cbo_Material.text), 50)) + " " + Trim(Left(UCase(cbo_Caracteristica.text), 80)) + " " + Trim(Left(UCase(cbo_Medida.text), 80)) + " " + Trim(Left(UCase(cbo_color.text), 80))

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Coneccion.Close
End Sub

Private Sub Cargar_Caracteristica()
'   carga caracteristica
    SQL = "select * from caracteristica_a where codfam ='" & Trim(Right(cbo_Familia.text, 2)) & "' OR descripcion is null order by codcar,codfam"
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    cbo_Caracteristica.Clear
    While Not RsAvios.EOF
      If IsNull(RsAvios!descripcion) Then
            cbo_Caracteristica.AddItem " " + Space(100) + RsAvios!Codcar
        Else
            cbo_Caracteristica.AddItem RsAvios!descripcion + Space(100) + RsAvios!Codcar
    End If
      
        RsAvios.MoveNext
    Wend
    'cbo_Caracteristica.AddItem "<Nuevo>"
    RsAvios.Close

End Sub


Private Sub Cargar_Medida()
'   carga medida
    SQL = "select * from Medida_a where codfam ='" & Right(cbo_Familia.text, 2) & "' OR descripcion is null order by codmed,codfam"
    RsAvios.Open SQL, VGcnx, adOpenStatic
    'Set RsAvios = Coneccion.Execute(SQL)
    cbo_Medida.Clear
    While Not RsAvios.EOF
       If IsNull(RsAvios!descripcion) Then
            cbo_Medida.AddItem " " + Space(100) + RsAvios!Codmed
        Else
            cbo_Medida.AddItem RsAvios!descripcion + Space(100) + RsAvios!Codmed
        End If
        RsAvios.MoveNext
    Wend
    RsAvios.Close
End Sub

