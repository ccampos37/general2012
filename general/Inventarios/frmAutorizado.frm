VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmAutorizado 
   Caption         =   "Personal Autorizado"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   Icon            =   "frmAutorizado.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   7
      Top             =   2280
      Width           =   5415
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Reporte"
         Height          =   675
         Left            =   3600
         Picture         =   "frmAutorizado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Nuevo"
         Height          =   675
         Index           =   0
         Left            =   240
         Picture         =   "frmAutorizado.frx":0D0C
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Grabar"
         Height          =   675
         Index           =   1
         Left            =   1080
         Picture         =   "frmAutorizado.frx":114E
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "E&ditar"
         Height          =   675
         Index           =   2
         Left            =   1920
         Picture         =   "frmAutorizado.frx":1590
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Caption         =   "&Eliminar"
         Height          =   675
         Index           =   3
         Left            =   2760
         Picture         =   "frmAutorizado.frx":19D2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   775
      End
      Begin VB.CommandButton cmdBotones 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   675
         Index           =   5
         Left            =   4440
         Picture         =   "frmAutorizado.frx":1E14
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   775
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   120
         Top             =   -120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin VB.Frame Frame0 
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5655
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   1815
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   3201
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1800
         MaxLength       =   20
         TabIndex        =   4
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   1800
         MaxLength       =   3
         TabIndex        =   3
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Personal Autorizado"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmAutorizado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim codayu As String
Dim Adoreg1 As ADODB.Recordset
Dim AdoReg2 As ADODB.Recordset
Dim cCad As String
Dim RegActual As Integer
Dim nFra As Integer
Dim nTipo As Integer
Dim nI As Integer

Private Sub CmdImprimir_Click()
    Dim CADENA As String
    Dim cNomRepor  As String

cNomRepor = "autorizado.RPT"
If Trim(cNomRepor) <> "" Then
    CrystalReport1.Reset
    CrystalReport1.WindowTitle = "Reporte de Personal Autorizado"
    CrystalReport1.ReportFileName = VGParamSistem.RutaReport + "\" + cNomRepor
 
                        
    CrystalReport1.Connect = VGcadenareport2
    CrystalReport1.StoredProcParam(0) = VGParamSistem.BDEmpresa
    CrystalReport1.StoredProcParam(1) = "12"
    
    CrystalReport1.formulas(0) = "emp ='" & VGparametros.RucEmpresa & "'"
    
    CrystalReport1.Destination = crptToWindow
    CrystalReport1.WindowShowPrintBtn = True
    CrystalReport1.WindowShowRefreshBtn = True
    CrystalReport1.WindowShowSearchBtn = True
    CrystalReport1.WindowShowPrintSetupBtn = True
    CrystalReport1.WindowState = crptMaximized
    If CrystalReport1.Status <> 2 Then
        CrystalReport1.Action = 1
    End If
Else
    MsgBox "No existe el nombre del Reporte, verifique en Formatos", vbInformation, "Información"
    Exit Sub
End If

End Sub

Private Sub Form_Load()
central Me
ADOConectar
Init_ControlDataGrid DataGrid1
If Adoreg1.RecordCount = 0 Then
    Botones_Init True
Else
    Botones_Set True
End If
SetDataGrid
Frame1.Visible = False
Frame0.Visible = True
nFra = 0
nTipo = 1
End Sub

Private Sub ADOConectar()
Set Adoreg1 = New ADODB.Recordset
Set AdoReg2 = New ADODB.Recordset

'Adoreg1.CursorType = adOpenDynamic           NO USAR CAMBIAR BOOKMARK

codayu = "12"
Adoreg1.Open "Select TCLAVE,TDESCRI from TABAYU  where TCOD= '" & codayu & "'  ", VGCNx, adOpenStatic
AdoReg2.Open "Select * from  TABAYU where TCOD = '" & codayu & "'", VGCNx, adOpenDynamic, adLockOptimistic
Set DataGrid1.DataSource = Adoreg1
End Sub

Public Sub SetDataGrid()
 DataGrid1.Refresh
 DataGrid1.Columns(0).Caption = "           Código"
 DataGrid1.Columns(1).Caption = "                            Nombre"
 DataGrid1.Columns(0).Width = 1500
 DataGrid1.Columns(1).Width = 3700
End Sub

Private Sub Cmdbotones_Click(Index As Integer)
Dim nIi As Integer
Dim tempi As Integer
Dim temps As String
Select Case Index
Case 0: 'Nuevo
         nTipo = 1
         nFra = 1
         Dim otext As TextBox
         For Each otext In Me.Text1
                otext.text = ""
         Next
         Frame1.Caption = "Nuevo Usuario"
         Frame1.Visible = True
         Frame0.Visible = False
         
         Botones_Set False
         Text1(0).Enabled = True
         Text1(0).SetFocus
 Case 1: 'Grabar
         Screen.MousePointer = 11
         If nTipo = 1 Then
            Dim flag As Boolean
            flag = False
            'buscar igual codigo
            With Adoreg1
                If .RecordCount <> 0 Then
                    .MoveFirst
                    .Find "tclave= '" & UCase(Trim(Text1(0).text)) & "'"
                    If Not .EOF Then
                        flag = True
                        Text1(0).text = ""
                        MsgBox "El Codigo  ya existe:  Ingrese de nuevo", vbInformation, "Ingreso de Datos"
                        Text1(0).SetFocus
                    End If
                End If
            End With
            If Not flag Then
                If Text1(1).text = "" Then
                    MsgBox "Ud. No ha ingresado el nombre del autorizado", vbInformation, "Ingreso de Datos"
                    Text1(1).SetFocus
                Else
                    'pasa
                    AdoReg2.AddNew
                    AdoReg2.Fields("Tcod") = "12"
                    AdoReg2.Fields("Tclave") = UCase(Trim(Text1(0).text))
                    AdoReg2.Fields("tdescri") = UCase(Trim(Text1(1).text))
                    'AdoReg2.Fields("TRESTA") = 0
                    'AdoReg2.Fields("TADVALOR") = 0
                    AdoReg2.Update
                    Adoreg1.Requery
                    Frame0.Visible = True
                    nFra = 0
                    Botones_Set True
                End If
            End If
        End If
         
        If nTipo = 2 Then
            AdoReg2.Fields("tdescri") = UCase(Trim(Text1(1).text))
            AdoReg2.UpdateBatch
            Adoreg1.Requery
            AdoReg2.Requery
       
            Frame1.Visible = False
            Frame0.Visible = True
            nFra = 0
            Botones_Set True
         End If
          Frame1.Visible = False
          Frame0.Visible = True
         SetDataGrid
         Screen.MousePointer = 1
         
 Case 2: 'Editar
        If Adoreg1.RecordCount > 0 Then
         If Adoreg1.Bookmark Then
            Screen.MousePointer = 11
            nTipo = 2
            AdoReg2.Bookmark = DataGrid1.Bookmark   '.Bookmark
'            AdoReg2.Delete adAffectCurrent
            nFra = 2
            Dim OTEXT1 As TextBox
            For Each OTEXT1 In Me.Text1
                OTEXT1.text = ""
            Next
            Text1(0).text = AdoReg2.Fields("TCLAVE")
            If Not IsNull(AdoReg2.Fields("TDESCRI")) Then Text1(1).text = AdoReg2.Fields("TDESCRI")
            Frame1.Caption = "Modificar Usuario"
            Frame1.Visible = True
            Frame0.Visible = False
            Botones_Set False
            Text1(0).Enabled = False
            Text1(1).SetFocus
            Screen.MousePointer = 1
         Else
            MsgBox "Debe seleccionar un Registro para editarlo", vbInformation
            Botones_Set False
            Cmdbotones_Click 5
         End If
         End If
 Case 3: 'Eliminar
          Dim op As Integer
          If AdoReg2.RecordCount > 0 Then
          op = MsgBox("Esta Seguro que desea Eliminar el registro actual ?", vbYesNo, "Eliminación de Registro")
          If op = vbYes Then
            AdoReg2.Bookmark = DataGrid1.Bookmark ' AdoReg2.Bookmark
            AdoReg2.Delete
            AdoReg2.Update
            'AdoReg2.UpdateBatch
            AdoReg2.Requery
            Adoreg1.Requery
            If Adoreg1.RecordCount = 0 Then
                Botones_Init True
            Else
                Botones_Set True
            End If
          End If
          SetDataGrid
          End If
 Case 5: 'Salir , Cancelar
         If Cmdbotones(5).Caption = "&Salir" Then
            Unload Me
         Else
            Cmdbotones(5).Caption = "&Salir"
            Frame1.Visible = False
            Frame0.Visible = True
            nFra = 0
            If Adoreg1.RecordCount = 0 Then
                Botones_Init True
            Else
                Botones_Set True
            End If
         End If
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
With Text1(Index)
     .SelStart = 0
    .SelLength = Len(.text)
End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index = 0 Then
        With Adoreg1
            If .RecordCount <> 0 Then
                .MoveFirst
                .Find "tclave= '" & UCase(Trim(Text1(0).text)) & "'"
                If Not .EOF Then
                    Text1(0).text = ""
                    MsgBox "El codigo  ya existe:  Ingrese de nuevo", vbInformation, "Ingreso de Datos"
                    Text1(0).SetFocus: Exit Sub
                End If
            End If
        End With
    End If
            
    SendKeys "{tab}"
    KeyAscii = 0
End If
End Sub
Public Sub Botones_Set(flag As Boolean)
'flag=false Nuevo; flag=true .etc...
 Cmdbotones(0).Enabled = flag 'Nuevo
 Cmdbotones(1).Enabled = Not flag 'Grabar
 Cmdbotones(2).Enabled = flag 'Editar
 Cmdbotones(3).Enabled = flag 'Eliminar
 If flag Then
  Cmdbotones(5).Caption = "&Salir" 'Salir
 Else
  Cmdbotones(5).Caption = "&Cancelar"
 End If
End Sub
Public Sub Botones_Init(flag As Boolean)
'flag=false Nuevo; flag=true .etc...
 Cmdbotones(0).Enabled = flag 'Nuevo
 Cmdbotones(1).Enabled = Not flag 'Grabar
 Cmdbotones(2).Enabled = Not flag 'Editar
 Cmdbotones(3).Enabled = Not flag 'Eliminar
 Cmdbotones(5).Caption = "&Salir" 'Salir
End Sub
Private Sub DataGrid1_Click()
RegActual = IIf(IsNull(DataGrid1.Bookmark), 0, DataGrid1.Bookmark)
End Sub
