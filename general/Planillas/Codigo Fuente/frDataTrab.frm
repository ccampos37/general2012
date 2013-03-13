VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frDataTrab 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Informativos"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frDataTrab.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      Height          =   360
      Left            =   3345
      TabIndex        =   3
      Top             =   4710
      Width           =   1170
   End
   Begin VB.CommandButton cmEliminar 
      Caption         =   "&Eliminar"
      Height          =   360
      Left            =   1785
      TabIndex        =   2
      Top             =   4710
      Width           =   1170
   End
   Begin VB.CommandButton cmagragar 
      Caption         =   "&Agregar"
      Height          =   360
      Left            =   225
      TabIndex        =   1
      Top             =   4710
      Width           =   1170
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3705
      Left            =   180
      TabIndex        =   0
      Top             =   885
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   6535
      _Version        =   393216
      AllowUpdate     =   -1  'True
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
      Caption         =   "Datos Informativos de Trabajadores"
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
         MarqueeStyle    =   4
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frDataTrab.frx":08CA
      ForeColor       =   &H8000000E&
      Height          =   645
      Left            =   915
      TabIndex        =   4
      Top             =   60
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frDataTrab.frx":0963
      Top             =   30
      Width           =   480
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   360
      Picture         =   "frDataTrab.frx":122D
      Top             =   210
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      Height          =   720
      Left            =   15
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frDataTrab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSDATA As New ADODB.Recordset

Private Sub CMAGRAGAR_CLICK()
    frAddColTrab.Show 1
    RSDATA.Requery
    Set xData.DataSource = RSDATA
End Sub

Private Sub CMELIMINAR_CLICK()
    If RSDATA.EOF Then Exit Sub
    If RSDATA!CODDATA = "BASICO2" Then
        MsgBox "Estas Campo es exclusivo del sistema, no se puede eliminar", vbExclamation
        Exit Sub
    End If
    
    MsgBox "Advertencia" & Chr(13) & Chr(10) & "La eliminacion de un registro de este tipo puede ocasionar - Si no esta seguro de sus funciones - la colision del sistema en el momento del calculo de planillas de remuneraciones", vbInformation
    If MsgBox("Realmente desea eliminar este registro. Advertencia! La informacion registrada se eliminara sin posibilidad de recuperacion", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    On Error GoTo ERRELIM
    DBSYSTEM.Execute "ALTER TABLE TRABAJADORES DROP COLUMN " & RSDATA!CODDATA
    DBSYSTEM.Execute "DELETE FROM DATATRAB WHERE CODDATA='" & RSDATA!CODDATA & "'"
    RSDATA.Requery
    Set xData.DataSource = RSDATA
    xData.Columns("CODDATA").Locked = True
    xData.Columns("TIPODATA").Locked = True
    Exit Sub
ERRELIM:
    MsgBox "No se puede actualizar el sistema de base de datos de trabajadores, Pues esta siendo utilizado por otro usuario", vbInformation
    Exit Sub
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    If Not ExisteTabla("DATATRAB") Then
        DBSYSTEM.Execute "CREATE TABLE DATATRAB (CODDATA TEXT(15),DESCDATA TEXT(30),TIPODATA TEXT(1))"
        MsgBox "El sistema de Plamillas ha actualizado su sistema", vbInformation
    End If
    RSDATA.Open "DATATRAB", DBSYSTEM, adOpenStatic, adLockOptimistic
    Set xData.DataSource = RSDATA
    xData.Columns("CODDATA").Locked = True
    xData.Columns("TIPODATA").Locked = True
End Sub

Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSDATA = Nothing
End Sub

