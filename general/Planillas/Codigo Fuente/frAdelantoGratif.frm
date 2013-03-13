VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{C5FF36B7-A67E-11D3-9D9D-E6F193F7F854}#9.0#0"; "ApliCTxt.ocx"
Begin VB.Form frAdelantoGratif 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adelanto de Gratificación"
   ClientHeight    =   5550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7215
   Icon            =   "frAdelantoGratif.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5550
   ScaleWidth      =   7215
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmRedondear 
      Caption         =   "&Redondear"
      Height          =   330
      Left            =   5565
      TabIndex        =   15
      Top             =   5130
      Width           =   1170
   End
   Begin AplisetControlText.Aplitext xPorcentaje 
      Height          =   315
      Left            =   240
      TabIndex        =   13
      Top             =   4695
      Width           =   570
      _ExtentX        =   1005
      _ExtentY        =   556
      Text            =   "85"
      TipoDato        =   "N"
   End
   Begin VB.CommandButton cmCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   330
      Left            =   2190
      TabIndex        =   8
      Top             =   5115
      Width           =   1230
   End
   Begin VB.CommandButton cmGrabar 
      Caption         =   "&Grabar"
      Height          =   330
      Left            =   2190
      TabIndex        =   7
      Top             =   4695
      Width           =   1230
   End
   Begin VB.CommandButton cmQuitar 
      Caption         =   "&Quitar Trab."
      Height          =   330
      Left            =   870
      TabIndex        =   6
      Top             =   5115
      Width           =   1230
   End
   Begin VB.CommandButton cmAdelanto 
      Caption         =   "% de Adelanto"
      Height          =   330
      Left            =   870
      TabIndex        =   5
      Top             =   4695
      Width           =   1230
   End
   Begin MSDataGridLib.DataGrid xData 
      Height          =   3435
      Left            =   240
      TabIndex        =   1
      Top             =   1170
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   6059
      _Version        =   393216
      HeadLines       =   2
      RowHeight       =   15
      FormatLocked    =   -1  'True
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
      Caption         =   "Adelanto de Gratificación Navidad"
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "CodTrab"
         Caption         =   "Codigo"
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
         DataField       =   "Nombres"
         Caption         =   "Apellidos y Nombres"
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
      BeginProperty Column02 
         DataField       =   "ImporteGrati"
         Caption         =   "Total Gratificación"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "Adelanto"
         Caption         =   "Adelanto de Gratif."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   1
            Format          =   "0.00 "
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   2
         BeginProperty Column00 
            Locked          =   -1  'True
            ColumnWidth     =   794.835
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   3000.189
         EndProperty
         BeginProperty Column02 
            Alignment       =   1
            Locked          =   -1  'True
            ColumnWidth     =   1200.189
         EndProperty
         BeginProperty Column03 
            Alignment       =   1
            ColumnWidth     =   1094.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración del Adelanto"
      Height          =   945
      Left            =   120
      TabIndex        =   0
      Top             =   90
      Width           =   5055
      Begin AplisetControlText.Aplitext xPeriodo 
         Height          =   300
         Left            =   1845
         TabIndex        =   14
         Top             =   420
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   529
         Locked          =   -1  'True
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cargar en Adelanto de:"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   450
         Width           =   1635
      End
   End
   Begin VB.Label xNumTrabs 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajadores: 0"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3570
      TabIndex        =   12
      Top             =   5190
      Width           =   1110
   End
   Begin VB.Label tot2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00 "
      Height          =   300
      Left            =   5565
      TabIndex        =   11
      Top             =   4710
      Width           =   1170
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Totales"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   3570
      TabIndex        =   10
      Top             =   4763
      Width           =   525
   End
   Begin VB.Label Tot1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
      Height          =   300
      Left            =   4365
      TabIndex        =   9
      Top             =   4710
      Width           =   1170
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Gratificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5535
      TabIndex        =   4
      Top             =   555
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Adelanto de"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   5535
      TabIndex        =   3
      Top             =   315
      Width           =   1035
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   6630
      Picture         =   "frAdelantoGratif.frx":030A
      Top             =   285
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      Height          =   4425
      Left            =   120
      Top             =   1095
      Width           =   6990
   End
End
Attribute VB_Name = "frAdelantoGratif"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSGRATI As New ADODB.Recordset
Private Sub CMADELANTO_CLICK()
    If Val(xPorcentaje.Text) = 0 Then
        MsgBox "El valor no puede ser cero", vbInformation
        xPorcentaje.SetFocus
        Exit Sub
    End If
    If Val(xPorcentaje.Text) > 100 Then
        MsgBox "No puede ser mayor a 100%", vbInformation
        xPorcentaje.SetFocus
        Exit Sub
    End If
    DBSYSTEM.Execute "UPDATE  [##ADELGRATI" & VGL_COMPUTER & "]  SET ADELANTO=IMPORTEGRATI*" & xPorcentaje.Text & "/100"
    REFRESCAR
End Sub

Private Sub CMCANCELAR_CLICK()
    Unload Me
End Sub

Private Sub CMGRABAR_CLICK()
    If MsgBox("Realmente desea grabar la información editada a los Adelantos de Gratificaciones", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    Dim MES As String
    If xPeriodo.Tag = "" Then
        MsgBox "No ha seleccionado un periodo donde cargarán los datos del adelanto de gratificación en proceso", vbInformation
        xPeriodo.SetFocus
        Exit Sub
    End If
    Dim XRC As Integer
    DBSYSTEM.Execute "UPDATE ADEL2000 SET ORIGEN=ORIGEN WHERE ORIGEN=" & xPeriodo.Tag, XRC
    If XRC > 0 Then
        XRC = MsgBox("Este grupo de adelantos de Gratificación ya contiene " & XRC & " adelantos. Desea ELIMINAR estos adelantos para poder almacenar los actuales. IMPORTANTE: NO EXISTE POSIBILIDAD DE SER RECUPERADOS", vbYesNoCancel)
        If XRC = vbCancel Then
            Exit Sub
        Else
            If XRC = vbYes Then
                CambiaPanelBD True
                DBSYSTEM.Execute "DELETE FROM ADEL2000 WHERE CODIGO=" & XRC
                CambiaPanelBD True
            Else
                MsgBox "IMPORTANTE: Si graba los adelantos procesados en la presente ventana se podrian duplicar los monto en caso cada trabajador ya tenga registrados otros adelantos de gratificación", vbInformation
            End If
        End If
    End If
    MES = DateSQL(DevuelveValor("SELECT MES FROM NOMBOL WHERE CODIGO=" & xPeriodo.Tag, DBSYSTEM))
    CambiaPanelBD True
    DBSYSTEM.Execute "INSERT INTO ADEL2000 (CODTRAB,MES,FECHAING,MONTO,NUMBOL,NOMBOL,ORIGEN) SELECT CODTRAB," & MES & " AS MESNOM," & DateSQL(Date) & " AS FECHAING,ADELANTO,0 AS T1, 0 AS T2," & xPeriodo.Tag & " FROM  [##ADELGRATI" & VGL_COMPUTER & "] "
    CambiaPanelBD False
    MsgBox "Información Grabada Satisfactoriamente. Se abandonará la presente ventana. Para ver los resultados, e imprimir "
    Unload Me
End Sub
Private Sub CMQUITAR_CLICK()
    If RSGRATI.RecordCount = 1 Then
        MsgBox "Se ha llegado al mínimo de registros para este proceso. Mejor Cancele la acción"
        cmGrabar.Enabled = False
        Exit Sub
    End If
    If MsgBox("Realmente desea eliminar el Adelanto de Gratificación del Trabajador " & RSGRATI!NOMBRES, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    If RSGRATI.EOF Then Exit Sub
    RSGRATI.Delete
    RSGRATI.MoveFirst
    CALCULATOTALES
End Sub

Private Sub CMREDONDEAR_CLICK()
    If RSGRATI.RecordCount = 0 Then Exit Sub
    RSGRATI.MoveFirst
    Do While Not RSGRATI.EOF
        RSGRATI!ADELANTO = Round(RSGRATI!ADELANTO, 0)
        RSGRATI.MoveNext
    Loop
    RSGRATI.MoveFirst
End Sub

Private Sub Form_Load()
    If ExisteTablaAux(" [##ADELGRATI" & VGL_COMPUTER & "] ") Then DBSYSTEM.Execute "DROP TABLE  [##ADELGRATI" & VGL_COMPUTER & "] "
    DBSYSTEM.Execute "CREATE TABLE  [##ADELGRATI" & VGL_COMPUTER & "]  (CODTRAB VARCHAR(8), NOMBRES varchar(35), IMPORTEGRATI  Numeric(20,2) , ADELANTO  Numeric(20,2) )"
    DBSYSTEM.Execute "INSERT INTO  [##ADELGRATI" & VGL_COMPUTER & "]  SELECT CODTRAB, NOMBRES, IMPORTEGRATI, 0 AS ADELANTO FROM PLANGRATI WHERE CODIGO=" & VPTRASPRM
    RSGRATI.Open " [##ADELGRATI" & VGL_COMPUTER & "] ", DBSYSTEM, adOpenKeyset, adLockOptimistic
    REFRESCAR
End Sub
Private Sub FORM_UNLOAD(CANCEL As Integer)
    Set RSGRATI = Nothing
End Sub
Private Sub IMAGE1_CLICK()
    MsgBox "ENTERPRISE SOLUTIONS S.A."
End Sub
Public Sub CALCULATOTALES()
    Tot1.Caption = Format(DevuelveValor("SELECT SUM(IMPORTEGRATI) AS T1 FROM  [##ADELGRATI" & VGL_COMPUTER & "] ", DBSYSTEM), "0.00 ")
    tot2.Caption = Format(DevuelveValor("SELECT SUM(ADELANTO) AS T1 FROM  [##ADELGRATI" & VGL_COMPUTER & "] ", DBSYSTEM), "0.00 ")
    xNumTrabs.Caption = "Trabajadores : " & RSGRATI.RecordCount
End Sub
Public Sub REFRESCAR()
    RSGRATI.Requery
    Set xData.DataSource = RSGRATI
    CALCULATOTALES
End Sub
Private Sub XDATA_AFTERCOLUPDATE(ByVal COLINDEX As Integer)
    RSGRATI.MOVE 0
    CALCULATOTALES
End Sub
Private Sub XPERIODO_DBLCLICK()
    Dim RSAUX As New ADODB.Recordset
    RSAUX.Open "SELECT NOMBOL.CODIGO, NOMBOL.NOMBRE FROM NOMBOL, MESESACT WHERE NOMBOL.MES=MESESACT.MESACTIVO AND DARADELANTO=1", DBSYSTEM, adOpenStatic, adLockReadOnly
    If RSAUX.RecordCount = 0 Or RSAUX.EOF Then
        MsgBox "No se han encontrado Meses Activos", vbInformation
        cmGrabar.Enabled = False
        Set RSAUX = Nothing
        Exit Sub
    End If
    frmComun.CONECTAR RSAUX
    frmComun.Show 1
    If VGUTIL(1) <> "" Then
        xPeriodo.Tag = VGUTIL(1)
        xPeriodo.Text = VGUTIL(2)
        Dim XRC As Integer
        DBSYSTEM.Execute "UPDATE ADEL2000 SET ORIGEN=ORIGEN WHERE ORIGEN=" & xPeriodo.Tag, XRC
        If XRC > 0 Then
            XRC = MsgBox("Este grupo de adelantos de Gratificación ya contiene " & XRC & " adelantos. Desea ELIMINAR estos adelantos para poder almacenar los actuales. IMPORTANTE: NO EXISTE POSIBILIDAD DE SER RECUPERADOS", vbYesNoCancel)
            If XRC = vbCancel Then
                xPeriodo.Tag = ""
                xPeriodo.Text = ""
            Else
                If XRC = vbYes Then
                    CambiaPanelBD True
                    DBSYSTEM.Execute "DELETE FROM ADEL2000 WHERE CODIGO=" & XRC
                    CambiaPanelBD True
                Else
                    MsgBox "IMPORTANTE: Si graba los adelantos procesados en la presente ventana se podrian duplicar los monto en caso cada trabajador ya tenga registrados otros adelantos de gratificación", vbInformation
                End If
            End If
        End If
    End If
    Set RSAUX = Nothing
End Sub

