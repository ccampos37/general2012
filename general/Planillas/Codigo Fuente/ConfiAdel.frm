VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "APLICTXT.OCX"
Begin VB.Form ConfiAdel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion de Adelantos"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7020
   Icon            =   "ConfiAdel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   7020
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab SSTab1 
      Height          =   5595
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   9869
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Adelantos"
      TabPicture(0)   =   "ConfiAdel.frx":08CA
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame2"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "Contabilidad"
      TabPicture(1)   =   "ConfiAdel.frx":08E6
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblConcepto"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Frame4"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command3"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command4"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).ControlCount=   7
      Begin VB.CommandButton Command4 
         Caption         =   "Cancelar"
         Height          =   255
         Left            =   5265
         TabIndex        =   21
         Top             =   5265
         Width           =   1470
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Aceptar"
         Height          =   255
         Left            =   3705
         TabIndex        =   20
         Top             =   5265
         Width           =   1440
      End
      Begin VB.Frame Frame2 
         Height          =   3960
         Left            =   -74940
         TabIndex        =   13
         Top             =   315
         Width           =   6900
         Begin VB.CommandButton cmCerrar 
            Cancel          =   -1  'True
            Caption         =   "&Salir"
            Height          =   315
            Left            =   5715
            TabIndex        =   18
            Top             =   3435
            Width           =   1020
         End
         Begin VB.CommandButton cmBorrar 
            Caption         =   "&Borrar"
            Height          =   315
            Left            =   2460
            TabIndex        =   17
            Top             =   3450
            Width           =   1020
         End
         Begin VB.CommandButton cmEditar 
            Caption         =   "&Editar"
            Height          =   315
            Left            =   1305
            TabIndex        =   16
            Top             =   3450
            Visible         =   0   'False
            Width           =   1020
         End
         Begin VB.CommandButton cmNuevo 
            Caption         =   "&Nuevo"
            Height          =   315
            Left            =   150
            TabIndex        =   15
            Top             =   3450
            Width           =   1020
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Configurar Cta Contable"
            Height          =   330
            Left            =   3600
            TabIndex        =   14
            Top             =   3435
            Width           =   1950
         End
         Begin MSDataGridLib.DataGrid xData 
            Height          =   3105
            Left            =   135
            TabIndex        =   19
            Top             =   195
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   5477
            _Version        =   393216
            AllowUpdate     =   0   'False
            Appearance      =   0
            HeadLines       =   2
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
                  LCID            =   10250
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
                  LCID            =   10250
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               MarqueeStyle    =   3
               AllowRowSizing  =   0   'False
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00808080&
            BackStyle       =   1  'Opaque
            Height          =   3705
            Left            =   90
            Top             =   150
            Width           =   6750
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Regresar>>"
         Height          =   270
         Left            =   105
         TabIndex        =   12
         Top             =   5280
         Width           =   1680
      End
      Begin VB.Frame Frame4 
         Caption         =   "Cuenta Haber"
         Height          =   2310
         Left            =   60
         TabIndex        =   10
         Top             =   2925
         Width           =   6675
         Begin MSDataGridLib.DataGrid XdataHaber 
            Height          =   1830
            Left            =   195
            TabIndex        =   11
            Top             =   315
            Width           =   6315
            _ExtentX        =   11139
            _ExtentY        =   3228
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            TabAction       =   2
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "SEC"
               Caption         =   "Sec"
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
               DataField       =   "CUENTA"
               Caption         =   "Cuenta"
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
            BeginProperty Column02 
               DataField       =   "TipAsi"
               Caption         =   "Tipo"
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
            BeginProperty Column03 
               DataField       =   "TIPASINOM"
               Caption         =   "Desc de tipo  de Asiento"
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
            BeginProperty Column04 
               DataField       =   "TIPOCTA"
               Caption         =   "TIPOCTA"
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
            BeginProperty Column05 
               DataField       =   "CONCEPT"
               Caption         =   "CONCEPT"
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
                  ColumnWidth     =   450.142
               EndProperty
               BeginProperty Column01 
                  Button          =   -1  'True
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column02 
                  Button          =   -1  'True
                  ColumnWidth     =   420.095
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2264.882
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Cuenta Debe"
         Height          =   2235
         Left            =   75
         TabIndex        =   8
         Top             =   630
         Width           =   6690
         Begin MSDataGridLib.DataGrid XdataDebe 
            Height          =   1830
            Left            =   180
            TabIndex        =   9
            Top             =   270
            Width           =   6345
            _ExtentX        =   11192
            _ExtentY        =   3228
            _Version        =   393216
            AllowUpdate     =   -1  'True
            HeadLines       =   1
            RowHeight       =   15
            FormatLocked    =   -1  'True
            AllowAddNew     =   -1  'True
            AllowDelete     =   -1  'True
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
            ColumnCount     =   6
            BeginProperty Column00 
               DataField       =   "SEC"
               Caption         =   "Sec"
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
               DataField       =   "CUENTA"
               Caption         =   "Cuenta"
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
            BeginProperty Column02 
               DataField       =   "TipAsi"
               Caption         =   "Tipo"
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
            BeginProperty Column03 
               DataField       =   "TIPASINOM"
               Caption         =   "Desc de tipo  de Asiento"
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
            BeginProperty Column04 
               DataField       =   "TIPOCTA"
               Caption         =   "TIPOCTA"
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
            BeginProperty Column05 
               DataField       =   "CONCEPT"
               Caption         =   "CONCEPT"
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
                  ColumnWidth     =   450.142
               EndProperty
               BeginProperty Column01 
                  Button          =   -1  'True
                  ColumnWidth     =   1454.74
               EndProperty
               BeginProperty Column02 
                  Button          =   -1  'True
                  ColumnWidth     =   420.095
               EndProperty
               BeginProperty Column03 
                  ColumnWidth     =   2264.882
               EndProperty
               BeginProperty Column04 
               EndProperty
               BeginProperty Column05 
               EndProperty
            EndProperty
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Descripción del concepto"
         Enabled         =   0   'False
         Height          =   1245
         Left            =   -74925
         TabIndex        =   1
         Top             =   4305
         Width           =   6855
         Begin VB.CommandButton cmCancelar 
            Caption         =   "&Cancelar"
            Height          =   315
            Left            =   5415
            TabIndex        =   4
            Top             =   810
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.CommandButton cmAceptar 
            Caption         =   "&Aceptar"
            Height          =   315
            Left            =   4245
            TabIndex        =   3
            Top             =   825
            Visible         =   0   'False
            Width           =   1080
         End
         Begin VB.ComboBox xTipo 
            Height          =   315
            ItemData        =   "ConfiAdel.frx":0902
            Left            =   2490
            List            =   "ConfiAdel.frx":090C
            TabIndex        =   2
            Text            =   "Combo1"
            Top             =   1410
            Visible         =   0   'False
            Width           =   2715
         End
         Begin AplisetControlText.Aplitext xNombre 
            Height          =   285
            Left            =   2505
            TabIndex        =   5
            Top             =   345
            Width           =   4035
            _ExtentX        =   7117
            _ExtentY        =   503
            Locked          =   -1  'True
            Text            =   ""
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Remuneración"
            Height          =   195
            Left            =   270
            TabIndex        =   7
            Top             =   1410
            Visible         =   0   'False
            Width           =   1620
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre del Concepto"
            Height          =   195
            Left            =   270
            TabIndex        =   6
            Top             =   405
            Width           =   1545
         End
      End
      Begin VB.Label lblConcepto 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Left            =   2925
         TabIndex        =   23
         Top             =   390
         Width           =   2610
      End
      Begin VB.Label Label2 
         Caption         =   "Concepto"
         Height          =   180
         Left            =   1920
         TabIndex        =   22
         Top             =   405
         Width           =   990
      End
   End
End
Attribute VB_Name = "ConfiAdel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSFORMS As ADODB.Recordset
Dim RSTIPDOC As New ADODB.Recordset
Dim WithEvents RSCTADEBE As ADODB.Recordset
Attribute RSCTADEBE.VB_VarHelpID = -1
Dim WithEvents RSCTAHABER As ADODB.Recordset
Attribute RSCTAHABER.VB_VarHelpID = -1
Dim FLAGMOV As Boolean


Private Sub CMACEPTAR_CLICK()
    Dim TIPOCONCEP As String
    If xNombre.Text = "" Then
        MsgBox "FALTA ESPECIFICAR UN NOMBRE", vbInformation
        xNombre.SetFocus
        Exit Sub
    End If
    TIPOCONCEP = DevuelveValor("SELECT  TIPO FROM CONCEPTOS WHERE CODIGO='" & xNombre.Tag & "'", DBSYSTEM)
    If Frame1.Tag = "NUEVO" Then
        DBSYSTEM.Execute "INSERT INTO CONFIADEL (CODIGO,NOMBRE,TIPO) VALUES ('" & xNombre.Tag & "','" & UCase(xNombre.Text) & "'," & TIPOCONCEP & ")"
    End If
    RSFORMS.Requery
    Set xData.DataSource = RSFORMS
    XDATA_ROWCOLCHANGE 0, 0
    FORMATEAR
    Frame1.Enabled = False
    cmAceptar.Visible = False
    cmCancelar.Visible = False
    OCULTAR
End Sub

Private Sub CMBORRAR_CLICK()
    If RSFORMS.EOF Or RSFORMS.BOF Or RSFORMS.RecordCount = 0 Then Exit Sub
    If MsgBox("ESTA SEGURO DE ELIMINAR EL CONCEPTO : " & RSFORMS!NOMBRE, vbYesNo + vbQuestion) = vbNo Then Exit Sub
    RSFORMS.Delete
    DBSYSTEM.Execute "DELETE FROM CTACONCEPTOQUIN WHERE CONCEPT='" & RSFORMS.Fields("CODIGO") & "'"
    Me.xNombre.Text = ""
    If RSFORMS.RecordCount > 0 Then RSFORMS.MoveFirst
    XDATA_ROWCOLCHANGE 0, 0
    FORMATEAR
End Sub

Private Sub CMCANCELAR_CLICK()
    Frame1.Enabled = False
    cmAceptar.Visible = False
    cmCancelar.Visible = False
    OCULTAR
End Sub

Private Sub cmCerrar_Click()
    Unload Me
End Sub

Private Sub CMEDITAR_CLICK()
    If RSFORMS.EOF Then Exit Sub
    Frame1.Tag = "EDITAR"
    Frame1.Enabled = True
    OCULTAR
    xNombre.SetFocus
End Sub

Private Sub CMNUEVO_CLICK()
    Frame1.Enabled = True
    Frame1.Tag = "NUEVO"
    OCULTAR
    xNombre.Text = ""
    xTipo.ListIndex = 0
    xNombre.SetFocus
End Sub

Private Sub Command1_Click()
    'frmHelpTmp.Show 1
    SSTab1.TabVisible(0) = True
    SSTab1.TabVisible(1) = False
    
End Sub

Private Sub Command2_Click()
    FLAGMOV = False
    Set RSCTADEBE = New ADODB.Recordset
    Set RSCTAHABER = New ADODB.Recordset
    RSCTADEBE.Open "SELECT * FROM CTACONCEPTOQUIN WHERE TIPOCTA='D' AND CONCEPT='" & RSFORMS.Fields("CODIGO") & "'", DBSYSTEM, adOpenDynamic, adLockBatchOptimistic
    RSCTAHABER.Open "SELECT * FROM CTACONCEPTOQUIN WHERE TIPOCTA='H' AND CONCEPT='" & RSFORMS.Fields("CODIGO") & "'", DBSYSTEM, adOpenDynamic, adLockBatchOptimistic
    Set XdataDebe.DataSource = RSCTADEBE
    Set XdataHaber.DataSource = RSCTAHABER
    FLAGMOV = True
    XdataHaber.Columns("TIPOCTA").Width = 0
    XdataHaber.Columns("CONCEPT").Width = 0
    XdataDebe.Columns("TIPOCTA").Width = 0
    XdataDebe.Columns("CONCEPT").Width = 0
    lblConcepto.Caption = RSFORMS.Fields("CODIGO")
    SSTab1.TabVisible(1) = True
    SSTab1.TabVisible(0) = False
End Sub

Private Sub Command3_Click()
On Error GoTo handler
    RSCTADEBE.UpdateBatch
    RSCTAHABER.UpdateBatch
    RSCTADEBE.Update
    RSCTAHABER.Update
    
Exit Sub
handler:
'MsgBox ERR.Description
Resume Next
End Sub

Private Sub Command4_Click()
Call Command1_Click
End Sub

Private Sub Form_Load()
    Set RSTIPDOC = New ADODB.Recordset
    RSTIPDOC.Open "SELECT * FROM CONCEPTOS WHERE ESESCRITO=1", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set RSFORMS = New ADODB.Recordset
    RSFORMS.Open "SELECT * FROM CONFIADEL", DBSYSTEM, adOpenKeyset, adLockOptimistic
    Set xData.DataSource = RSFORMS
    FORMATEAR
    xTipo.ListIndex = 0
    XDATA_ROWCOLCHANGE 0, 0
    
    SSTab1.Tab = 1
    SSTab1.TabVisible(1) = False
    SSTab1.TabVisible(0) = True
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSFORMS = Nothing
End Sub

Public Sub OCULTAR()
    If Frame1.Enabled Then
        cmAceptar.Visible = True
        cmCancelar.Visible = True
        cmBorrar.Enabled = False
        cmCerrar.Enabled = False
        cmEditar.Enabled = False
        cmNuevo.Enabled = False
        xData.Enabled = False
    Else
        cmAceptar.Visible = False
        cmCancelar.Visible = False
        cmBorrar.Enabled = True
        cmCerrar.Enabled = True
        cmEditar.Enabled = True
        cmNuevo.Enabled = True
        xData.Enabled = True
    End If
End Sub

Private Sub RSCTADEBE_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'NO SQL
    On Error GoTo ERRMOV
    If Not FLAGMOV Then Exit Sub
    XdataDebe.Columns("TIPOCTA") = "D"
    XdataDebe.Columns("CONCEPT").Value = Trim(lblConcepto.Caption)
    Exit Sub
ERRMOV:
    Exit Sub
End Sub

Private Sub RSCTAHABER_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)

On Error GoTo ERRMOV
    If Not FLAGMOV Then Exit Sub
    XdataHaber.Columns("TIPOCTA") = "H"
    XdataHaber.Columns("CONCEPT").Value = Trim(lblConcepto.Caption)
    Exit Sub
ERRMOV:
    Exit Sub

End Sub

Private Sub XDATA_ROWCOLCHANGE(LASTROW As Variant, ByVal LASTCOL As Integer)
    If RSFORMS.BOF Or RSFORMS.EOF Or RSFORMS.RecordCount = 0 Then Exit Sub
    If Frame1.Enabled Then Exit Sub
    xNombre.Text = RSFORMS!NOMBRE
    xNombre.Tag = RSFORMS!Codigo
End Sub

Private Sub XdataDebe_ButtonClick(ByVal ColIndex As Integer)
    Dim DESCAUX As String
    XdataHaber.SetFocus
    Screen.MousePointer = 13
    Select Case ColIndex
        Case 1
            If REGSISTEMA.scTieneStConta Then
                XdataDebe.Columns("CUENTA").Text = SELCUENTA(XdataDebe.Columns("CUENTA").Text)
            End If
        Case 2
            XdataDebe.Columns(2).Text = SELTIPOASIS(XdataDebe.Columns(2).Text, DESCAUX)
            If DESCAUX <> "" Then XdataDebe.Columns(3).Text = DESCAUX
    End Select
    XdataDebe.SetFocus
    Screen.MousePointer = 1
End Sub

Private Sub XdataHaber_ButtonClick(ByVal ColIndex As Integer)
    Dim DESCAUX As String
    Screen.MousePointer = 13
    XdataDebe.SetFocus
    Select Case ColIndex
        Case 1
            If REGSISTEMA.scTieneStConta Then
                XdataHaber.Columns("CUENTA").Text = SELCUENTA(XdataHaber.Columns("CUENTA").Text)
            End If
        Case 2
            XdataHaber.Columns(2).Text = SELTIPOASIS(XdataHaber.Columns(2).Text, DESCAUX)
            If DESCAUX <> "" Then XdataHaber.Columns(3).Text = DESCAUX
    End Select
    XdataHaber.SetFocus
    Screen.MousePointer = 1
End Sub
Private Sub XNOMBRE_DblClick()
Dim RSAUX As New ADODB.Recordset
    frmComun.CONECTAR RSTIPDOC
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xNombre.Text = VGUTIL(2)
        xNombre.Tag = VGUTIL(1)
          RSAUX.Open "Select TIPO from CONCEPTOS where CODIGO='" & xNombre.Tag & "'", DBSYSTEM, adOpenKeyset, adLockOptimistic
          If RSAUX.RecordCount > 0 Then
            xTipo.ListIndex = RSAUX.Fields(0) - 1
          End If
    End If
End Sub

Public Sub FORMATEAR()
    xData.Columns(0).Width = 1000
    xData.Columns(1).Width = 3500
    'xData.Columns(2).Width = 700
End Sub
Private Function SELCUENTA(TEXTO As String) As String
    Dim RSAUX As New ADODB.Recordset
    Set RSAUX = New ADODB.Recordset
    RSAUX.Open "SELECT PLANCTA_CODIGO,PLANCTA_DESCRIPCION FROM PLAN_CUENTA_NACIONAL WHERE PLANCTA_NIVEL=" & REGSISTEMA.scNivelCta, CONECTARDBSQL(REGSISTEMA.scRutaEmpresaWenco & "BDCONTABILIDAD"), adOpenKeyset, adLockReadOnly
    VGUTIL(1) = ""
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        SELCUENTA = VGUTIL(1)
      Else
        SELCUENTA = TEXTO
    End If
End Function
Private Function SELTIPOASIS(TEXTO As String, Optional ByRef DESC As String) As String
'NO SQL
    Dim RSTIP As New ADODB.Recordset
    Dim CAMPOS As Variant
    RSTIP.Fields.Append "COD", adInteger
    RSTIP.Fields.Append "DESC", adVarChar, 25
    CAMPOS = Array("COD", "DESC")
    RSTIP.Open
    RSTIP.AddNew CAMPOS, Array("1", "SIMPLE")
    RSTIP.AddNew CAMPOS, Array("2", "POR TRABAJADOR")
    RSTIP.AddNew CAMPOS, Array("3", "POR CENTRO DE COSTOS")
    RSTIP.AddNew CAMPOS, Array("4", "POR A.F.P.")
    RSTIP.AddNew CAMPOS, Array("5", "POR TRABAJADOR Y C.C.")
    VGUTIL(1) = ""
    RSTIP.Filter = "COD='1' OR COD='2' OR COD='3'"
    frmComun.CONECTAR RSTIP
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        SELTIPOASIS = VGUTIL(1)
        DESC = VGUTIL(2)
      Else
        SELTIPOASIS = TEXTO
        DESC = ""
    End If
End Function

Sub MOSTRAR_TAB(ByVal CONCEPTO As String)
    Set RSCTADEBE = New ADODB.Recordset
    Set RSCTAHABER = New ADODB.Recordset
    RSCTADEBE.Open "SELECT * FROM CTACONCEPTOQUIN WHERE TIPOCTA='D' AND CONCEPT='" & CONCEPTO & "'", DBSYSTEM, adOpenDynamic, adLockBatchOptimistic
    RSCTAHABER.Open "SELECT * FROM CTACONCEPTOQUIN WHERE TIPOCTA='H' AND CONCEPT='" & CONCEPTO & "'", DBSYSTEM, adOpenDynamic, adLockBatchOptimistic
    Set XdataDebe.DataSource = RSCTADEBE
    Set XdataHaber.DataSource = RSCTAHABER
End Sub
