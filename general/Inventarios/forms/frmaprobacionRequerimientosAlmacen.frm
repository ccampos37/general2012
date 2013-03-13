VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmAprobacionRequerimientosAlmacen 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Aprobacion Requerimientos"
   ClientHeight    =   5700
   ClientLeft      =   1710
   ClientTop       =   1710
   ClientWidth     =   9795
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmaprobacionRequerimientosAlmacen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   9795
   Begin MSComCtl2.DTPicker txtfec 
      Height          =   288
      Left            =   1560
      TabIndex        =   1
      Top             =   1704
      Width           =   1212
      _ExtentX        =   2143
      _ExtentY        =   503
      _Version        =   393216
      Format          =   62914561
      CurrentDate     =   37015
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   5184
      Picture         =   "frmaprobacionRequerimientosAlmacen.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4896
      Width           =   775
   End
   Begin VB.CommandButton cmdGra 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   3849
      Picture         =   "frmaprobacionRequerimientosAlmacen.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4896
      Width           =   775
   End
   Begin VB.Frame fraCabec 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   135
      TabIndex        =   11
      Top             =   0
      Width           =   9516
      Begin VB.TextBox txtNum 
         Height          =   285
         Left            =   1110
         MaxLength       =   13
         TabIndex        =   0
         Top             =   200
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número  :"
         Height          =   195
         Left            =   375
         TabIndex        =   14
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame fraDatos 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1665
      Left            =   144
      TabIndex        =   4
      Top             =   600
      Width           =   9516
      Begin VB.Label lblEst 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   7056
         TabIndex        =   19
         Top             =   720
         Width           =   396
      End
      Begin VB.Label lblEsta 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   7536
         TabIndex        =   18
         Top             =   720
         Width           =   1788
      End
      Begin VB.Label lblEnt 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   4560
         TabIndex        =   17
         Top             =   720
         Width           =   1092
      End
      Begin VB.Label lblEmi 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblPro 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Estado  :"
         Height          =   192
         Left            =   6216
         TabIndex        =   13
         Top             =   732
         Width           =   636
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Entrega   :"
         Height          =   192
         Left            =   3720
         TabIndex        =   12
         Top             =   732
         Width           =   732
      End
      Begin VB.Label lblRuc 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   7728
         TabIndex        =   10
         Top             =   360
         Width           =   1596
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C.  :"
         Height          =   192
         Left            =   7056
         TabIndex        =   9
         Top             =   372
         Width           =   612
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor     :"
         Height          =   195
         Left            =   375
         TabIndex        =   8
         Top             =   380
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emisión         :"
         Height          =   195
         Left            =   375
         TabIndex        =   7
         Top             =   740
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha           :"
         Height          =   195
         Left            =   375
         TabIndex        =   6
         Top             =   1215
         Width           =   990
      End
      Begin VB.Label lblProv 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   360
         Width           =   2970
      End
   End
   Begin VB.TextBox txtCol 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   5310
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   4110
      Visible         =   0   'False
      Width           =   225
   End
   Begin TrueOleDBGrid70.TDBGrid DBGrid1 
      Height          =   2505
      Left            =   150
      TabIndex        =   21
      Top             =   2310
      Width           =   9570
      _ExtentX        =   16880
      _ExtentY        =   4419
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=15,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
End
Attribute VB_Name = "frmAprobacionRequerimientosAlmacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rbusca1 As New ADODB.Recordset
Dim rsdeta2 As New ADODB.Recordset
Dim adodc1 As New ADODB.Recordset
Dim Adodc2 As New ADODB.Recordset
Dim Adodc3 As New ADODB.Recordset
Dim Conexion As String
Dim vgtipo As String
Dim txtTF As String
Dim SQL As String
Dim nTra As Integer
Dim Mensaje As String

Private Sub cmdGra_Click()
    Dim SQLc As String
    Dim SQLd As String
    Dim I As Integer, TipoIng As Integer
    Dim vNI As Integer
    Dim vNF As Single, vNC As Single
    Dim vNP As Single, vNP1 As Single
    On Error GoTo GrabErr
    
        If CDate(txtfec.Value) < CDate(lblEmi) Then
            Mensaje = "La Fecha debe ser igual o posterior que la Fecha de emisión"
            MsgBox Mensaje, vbExclamation, "Mensaje"
            txtfec.SetFocus
            Exit Sub
        End If
    
    TipoIng = Ingreso_Realizado
    If TipoIng = 0 Then
        Mensaje = "No se puede grabar." & vbCrLf & "No se ha recepcionado ningún artículo"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        Exit Sub
    End If

Mensaje = "¿Desea guardar los cambios realizados?"
If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
   vNI = 1

   nTra = 1
   VGCNx.BeginTrans
       
   SQLc = "UPDATE co_cabordcompra SET estadooccodigo='" & Format(TipoIng, "0") & _
        "' WHERE tipoordencodigo='" & vgtipo & "' and oc_cnumord='" & txtNum & "'"
   VGCNx.Execute SQLc
            
   If rsdeta2.RecordCount > 0 Then
      rsdeta2.MoveFirst
      Do Until rsdeta2.EOF
         If rsdeta2.Fields(6) > 0 Then
            SQLd = "UPDATE co_detordcompra SET oc_ncanten=" & rsdeta2.Fields(6) & ","
            SQLd = SQLd & "oc_nsaldo =" & rsdeta2.Fields(6) & ","
            If Val(rsdeta2.Fields(6)) >= Val(rsdeta2.Fields(5)) Then
               SQLd = SQLd & "oc_situacionorden='" & Format(2, "0") & "'"
             ElseIf Val(rsdeta2.Fields(6)) < Val(rsdeta2.Fields(5)) Then
               SQLd = SQLd & "oc_situacionorden='" & Format(1, "0") & "'"
            End If
            SQLd = SQLd & " WHERE tipoordencodigo='" & vgtipo & "' and oc_cnumord='" & txtNum & "' AND oc_citem ='" & rsdeta2.Fields(0) & "'"
            VGCNx.Execute SQLd
         End If
         rsdeta2.MoveNext
      Loop
   End If
       
   
   VGCNx.CommitTrans
   nTra = 0
  'adodc1.Requery
        
   Mensaje = "Se Aprobo el requerimiento"
   MsgBox Mensaje, vbInformation, "Ingreso"
   txtNum = ""
   txtNum.SetFocus
End If
Exit Sub

GrabErr:
    MsgBox Err.Description
   Resume
    If nTra = 1 Then VGCNx.RollbackTrans

End Sub

Private Sub CmdSalir_Click()
    Unload frmReferencia
    Unload frmingresoOC
    Unload Me
End Sub


Private Sub Form_Load()
    central Me
    'txtNum.SetFocus
End Sub

Sub Limpiar()
    lblPro = "": lblProv = "": lblRuc = ""
    lblEmi = "": lblEnt = "": lblEst = ""
    lblEsta = ""
    
End Sub

Private Sub txtFec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not ValidFecha(txtfec) Then
            Mensaje = "Fecha No Válida"
            MsgBox Mensaje, vbExclamation, "Error"
            txtfec.SetFocus
        Else
            If CDate(txtfec.Value) < CDate(lblEmi) Then
                Mensaje = "La Fecha debe ser igual o posterior que la Fecha de emisión"
                MsgBox Mensaje, vbExclamation, "Mensaje"
                txtfec.SetFocus
             End If
        End If
    End If
End Sub

Private Sub Txtnum_Change()
    If lblPro <> "" Then
        Limpiar
        fraDatos.Enabled = False
        cmdGra.Enabled = False
    End If
End Sub

Private Sub txtNum_DblClick()
    Set Adodc2 = New ADODB.Recordset
    SQL = "SELECT a.tipoordencodigo,a.oc_cnumord, a.oc_dfecdoc, b.estadoocdescripcion " & _
             " FROM co_cabordcompra a inner join co_estadorequerimiento b" & _
             " on a.estadooccodigo=b.estadooccodigo " & _
             " inner join co_tipodeorden c on a.tipoordencodigo=c.tipoordencodigo " & _
             " where b.estadooccodigo=0 and a.oc_estadoorden=0 and c.flagrequerimientos=1 "
    Adodc2.Open SQL, VGCNx, adOpenStatic, adLockReadOnly
    frmReferencia1.Conectar Adodc2, SQL
    frmReferencia1.Caption = "Requerimientos"
    frmReferencia1.inicio
    frmReferencia1.Show vbModal
    Adodc2.Close
    
    
frmReferencia.Conectar Adodc2, "Select ACODIGO, ADESCRI,AUNIDAD from MaeArt"
 
    If vGUtil(2) <> "" Then
        vgtipo = vGUtil(1)
        txtNum = vGUtil(2)
      '  lblSol = vGUtil(2)
      '  txtcen.SetFocus
    End If
End Sub

Private Sub txtNum_GotFocus()
    Set DBGrid1.DataSource = Nothing
    Enfoque txtNum
End Sub

Private Sub txtNum_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtNum_DblClick
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtNum <> "" Then
            txtNum = Format(txtNum, "00000000000")
            If Not Existe(1, txtNum, "co_cabordcompra", "oc_cnumord", False) Then
                Mensaje = "El Número de Orden de Compra ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Mensaje"
                Enfoque txtNum
                txtNum.SetFocus
                Exit Sub
            Else
                If Not Estado_Valido Then
                    Mensaje = "Estado de Orden de compra no válido"
                    MsgBox Mensaje, vbExclamation, "Mensaje"
                    Enfoque txtNum
                    txtNum.SetFocus
                    Exit Sub
                Else
                    Muestra_datos_de_co_cabordcompra
                    CargaGrilla
                    fraDatos.Enabled = True
                    cmdGra.Enabled = True
                    txtfec = VG_FecTrab
                    txtfec.SetFocus
                End If
            End If
        End If
    End If
    Enteros_Positivos KeyAscii, txtNum
End Sub

Public Function CargaGrilla()
Call cargar_grilla2
Set rbusca1 = Nothing

SQL = "SELECT oc_citem,oc_ccodigo,oc_cdesref,ord_fabnum,oc_ncantid,oc_nsaldo,oc_nsaldo as oc_ncanten " & _
      " FROM co_detordcompra WHERE oc_cnumord='" & txtNum & "' and oc_situacionorden<>'2'  ORDER BY oc_citem"
Set rbusca1 = VGCNx.Execute(SQL)
If rbusca1.RecordCount > 0 Then
   rbusca1.MoveFirst
   Do Until rbusca1.EOF
      rsdeta2.AddNew
      rsdeta2.Fields(0) = rbusca1!oc_citem
      rsdeta2.Fields(1) = rbusca1!oc_ccodigo
      rsdeta2.Fields(2) = ESNULO(Left(rbusca1!oc_cdesref, 30), "")
      rsdeta2.Fields(3) = rbusca1!ord_fabnum
      rsdeta2.Fields(4) = rbusca1!oc_ncantid
      rsdeta2.Fields(5) = rbusca1!oc_nsaldo
      rsdeta2.Fields(6) = rbusca1!oc_ncanten
      rbusca1.MoveNext
   Loop
End If
rbusca1.Close
Set rbusca1 = Nothing
End Function

Public Function cargar_grilla2()

Set rsdeta2 = Nothing
Call rsdeta2.Fields.Append("Item", adVarChar, 20)
Call rsdeta2.Fields.Append("Codigo", adVarChar, 20)
Call rsdeta2.Fields.Append("Descripcion", adVarChar, 30)
Call rsdeta2.Fields.Append("Ord.Fab.", adVarChar, 20)
Call rsdeta2.Fields.Append("Cant.Pedida", adVarChar, 20)
Call rsdeta2.Fields.Append("Saldo", adVarChar, 20)
Call rsdeta2.Fields.Append("Cant.Alm.", adVarChar, 20)
rsdeta2.Open

Set DBGrid1.DataSource = Nothing
Set DBGrid1.DataSource = rsdeta2

DBGrid1.Columns(0).AllowFocus = False
DBGrid1.Columns(1).AllowFocus = False
DBGrid1.Columns(2).AllowFocus = False
DBGrid1.Columns(3).AllowFocus = False
DBGrid1.Columns(4).AllowFocus = False
DBGrid1.Columns(5).AllowFocus = False
DBGrid1.Columns(0).Width = 400
DBGrid1.Columns(1).Width = 1200
DBGrid1.Columns(2).Width = 4000
DBGrid1.Columns(3).Width = 1000
DBGrid1.Columns(4).Width = 800
DBGrid1.Columns(5).Width = 800
DBGrid1.Columns(6).Width = 800
DBGrid1.Columns(4).NumberFormat = "###,##0.00"
DBGrid1.Columns(5).NumberFormat = "###,##0.00"
DBGrid1.Columns(6).NumberFormat = "###,##0.00"

DBGrid1.Refresh

End Function

Function Estado_Valido() As Boolean
    Dim vest As String
    
    vest = Devolver_Dato(1, txtNum, "co_cabordcompra", "oc_cnumord", False, "estadooccodigo")
    Estado_Valido = False
    If vest <> "2" Then Estado_Valido = True
End Function

Sub Muestra_datos_de_co_cabordcompra()
    Static adodc1 As New ADODB.Recordset
    Set adodc1 = New ADODB.Recordset
    SQL = "SELECT oc_ccodpro,oc_crazsoc=clienterazonsocial,oc_dfecdoc,oc_dfecent,estadooccodigo,oc_ccodmon,"
    SQL = SQL & "oc_csolict FROM co_cabordcompra A left  join cp_proveedor b "
    SQL = SQL & " on a.oc_ccodpro=b.clientecodigo WHERE oc_cnumord='" & txtNum & "'"
    adodc1.Open SQL, VGCNx, adOpenDynamic, adLockOptimistic
    
    lblPro = adodc1("oc_ccodpro")
    lblProv = ESNULO(adodc1("oc_crazsoc"), "")
    lblRuc = Devolver_Dato(1, lblPro, "cp_proveedor", "clientecodigo", False, "clienteruc")
    lblEmi = adodc1("oc_dfecdoc")
    lblEnt = adodc1("oc_dfecent")
    lblEst = adodc1("estadooccodigo")
    lblEsta = Devolver_Dato(1, lblEst, "co_estadorequerimiento", "estadooccodigo", False, "estadoocdescripcion")
End Sub

Function Ingreso_Realizado() As Integer
    Dim I As Integer
    Dim tSal As Single
    Dim tRec As Single
If rsdeta2.RecordCount > 0 Then
   rsdeta2.MoveFirst
   Do Until rsdeta2.EOF
      tSal = tSal + rsdeta2.Fields(5)
      tRec = tRec + rsdeta2.Fields(6)
      rsdeta2.MoveNext
   Loop
End If
If tRec = 0 Then
   Ingreso_Realizado = 0
    ElseIf tRec < tSal Then
        Ingreso_Realizado = 1
    Else
        Ingreso_Realizado = 2
    End If
End Function
