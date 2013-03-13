VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D2B97638-05A0-43C1-BDD0-A8D84599A1D6}#4.0#0"; "controlayuda.ocx"
Begin VB.Form frmemisionOCdetalle 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Artículo"
   ClientHeight    =   7212
   ClientLeft      =   1380
   ClientTop       =   1860
   ClientWidth     =   7104
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7212
   ScaleWidth      =   7104
   StartUpPosition =   2  'CenterScreen
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   1980
      Left            =   96
      TabIndex        =   40
      Top             =   4368
      Width           =   6924
      _ExtentX        =   12213
      _ExtentY        =   3493
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
      Splits(0).RecordSelectorWidth=   508
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=3048"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2963"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=3048"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2963"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=7.8,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      Caption         =   "Compras por Articulo"
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   13160660
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=780,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=780,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=780,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=780,.italic=0"
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
   Begin VB.Frame Frame1 
      Height          =   1500
      Left            =   96
      TabIndex        =   27
      Top             =   48
      Width           =   6876
      Begin VB.TextBox txtRef 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5412
         MultiLine       =   -1  'True
         TabIndex        =   31
         Top             =   1056
         Width           =   855
      End
      Begin VB.TextBox txtURe 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3336
         TabIndex        =   30
         Top             =   1080
         Width           =   735
      End
      Begin VB.TextBox txtCan 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         MultiLine       =   -1  'True
         TabIndex        =   29
         Top             =   636
         Width           =   1095
      End
      Begin ctrlayuda_f.Ctr_Ayuda CtrAyu_articulo 
         Height          =   348
         Left            =   1152
         TabIndex        =   28
         Top             =   240
         Width           =   5340
         _ExtentX        =   9419
         _ExtentY        =   614
         XcodMaxLongitud =   20
         xcodwith        =   1000
         NomTabla        =   "maeart"
         ListaCampos     =   "acodigo(1),adescri(1),acodigo2(2),aunidad(2)"
         XcodCampo       =   "acodigo"
         XListCampo      =   "adescri"
         ListaCamposDescrip=   "Vodigo,Descripcion"
         ListaCamposText =   "acodigo,adescri,acodigo2,aunidad"
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Referencia"
         Height          =   192
         Left            =   4512
         TabIndex        =   39
         Top             =   1164
         Width           =   780
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Unid. Ref."
         Height          =   192
         Left            =   2568
         TabIndex        =   38
         Top             =   1140
         Width           =   720
      End
      Begin VB.Label lblFab 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1152
         TabIndex        =   37
         Top             =   624
         Width           =   1452
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   192
         Left            =   144
         TabIndex        =   36
         Top             =   264
         Width           =   492
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cantidad"
         Height          =   192
         Left            =   4512
         TabIndex        =   35
         Top             =   696
         Width           =   636
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Unidad"
         Height          =   192
         Left            =   96
         TabIndex        =   34
         Top             =   1164
         Width           =   516
      End
      Begin VB.Label lblUnidad 
         AutoSize        =   -1  'True
         Caption         =   "Fabricante"
         Height          =   192
         Left            =   240
         TabIndex        =   33
         Top             =   672
         Width           =   756
      End
      Begin VB.Label lblUni 
         BorderStyle     =   1  'Fixed Single
         Height          =   288
         Left            =   1176
         TabIndex        =   32
         Top             =   1056
         Width           =   732
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   105
      TabIndex        =   22
      Top             =   3348
      Width           =   6948
      Begin VB.TextBox txtCo1 
         Height          =   285
         Left            =   1320
         TabIndex        =   24
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox txtordfab 
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   23
         Top             =   240
         Width           =   1740
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Comentario"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Orden fab."
         Height          =   192
         Left            =   288
         TabIndex        =   25
         Top             =   288
         Width           =   744
      End
   End
   Begin VB.Frame Frame3 
      Height          =   996
      Left            =   144
      TabIndex        =   12
      Top             =   2508
      Width           =   6900
      Begin VB.TextBox txtPIg 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5256
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   216
         Width           =   735
      End
      Begin VB.Label lblTNe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   288
         Left            =   5112
         TabIndex        =   20
         Top             =   576
         Width           =   1332
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   192
         Left            =   6096
         TabIndex        =   19
         Top             =   216
         Width           =   120
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Importe IGV."
         Height          =   192
         Left            =   312
         TabIndex        =   18
         Top             =   600
         Width           =   888
      End
      Begin VB.Label lblIgv 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   288
         Left            =   1272
         TabIndex        =   17
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label lblTCo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   285
         Left            =   1320
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Porcen. IGV."
         Height          =   192
         Left            =   4176
         TabIndex        =   15
         Top             =   216
         Width           =   912
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Total Neto"
         Height          =   192
         Left            =   4032
         TabIndex        =   14
         Top             =   576
         Width           =   756
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total Compra"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame2 
      Height          =   996
      Left            =   96
      TabIndex        =   5
      Top             =   1572
      Width           =   6948
      Begin VB.TextBox txtPDe 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   5208
         MaxLength       =   20
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   264
         Width           =   735
      End
      Begin VB.TextBox txtPUn 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         MaxLength       =   50
         MultiLine       =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblPNe 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   288
         Left            =   5160
         TabIndex        =   21
         Top             =   624
         Width           =   1332
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "%"
         Height          =   192
         Left            =   6048
         TabIndex        =   11
         Top             =   264
         Width           =   120
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Importe Dsct."
         Height          =   192
         Left            =   192
         TabIndex        =   10
         Top             =   600
         Width           =   948
      End
      Begin VB.Label lblDes 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         Height          =   288
         Left            =   1320
         TabIndex        =   9
         Top             =   600
         Width           =   1212
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Precio Unit."
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Porcen. Dsct."
         Height          =   192
         Left            =   4128
         TabIndex        =   7
         Top             =   264
         Width           =   972
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Precio Neto"
         Height          =   192
         Left            =   4080
         TabIndex        =   6
         Top             =   624
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   3504
      Picture         =   "frmEmisionOCdetalle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6480
      Width           =   775
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   1728
      Picture         =   "frmEmisionOCdetalle.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6480
      Width           =   775
   End
End
Attribute VB_Name = "frmemisionOCdetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public activado As Boolean
Public cancelado As Boolean
Public Igv As Single
Public Tipo As String
Dim Mensaje As String

Private Sub cmdCancel_Click()
    cancelado = True
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If Trim(CtrAyu_articulo.xclave) = "" Then
        Mensaje = "Debe ingresar Código de Artículo"
        MsgBox Mensaje, vbExclamation, "Error"
        CtrAyu_articulo.SetFocus
        Exit Sub
    End If
    
    If txtURe <> "" Then
        If Not txtRef.Enabled Then
            If Not Existe(1, txtURe, "tabunimed", "um_abrev", False) Then
                Mensaje = "Unidad de referencia no válida"
                MsgBox Mensaje, vbExclamation, "Error"
                txtURe.SetFocus
                Exit Sub
            Else
                txtURe_KeyPress 13
                cmdOK.SetFocus
            End If
        End If
        If Val(txtRef) = 0 Then
            Mensaje = "Debe especificar Orden de FabricacionccionReferencia"
            MsgBox Mensaje, vbExclamation, "Error"
            txtRef.SetFocus
            Exit Sub
        End If
    End If
    If Val(txtPUn) = 0 Then
        Mensaje = "Debe especificar Precio Unitario"
        MsgBox Mensaje, vbExclamation, "Error"
        txtPUn.SetFocus
        Exit Sub
    End If
    cancelado = False
    CtrAyu_articulo.Enabled = True
    CtrAyu_articulo.SetFocus
    Me.Hide
End Sub

Private Sub CtrAyu_articulo_AlDevolverDato(ByVal ColecCampos As ADODB.Fields)
    lblDes = CtrAyu_articulo.xnombre
    txtCan = "0.00"
    lblUni = ColecCampos("aunidad")
    txtCan.Enabled = True
End Sub

Private Sub Form_Activate()
    Igv = Val(txtPIg)
End Sub

Private Sub Form_Load()
Call CtrAyu_articulo.conexion(VGcnx)
End Sub

Private Sub txtCan_Change()
    If Val(txtCan) = 0 Then
        txtURe = ""
        txtURe.Enabled = False
        txtPUn = "0.00"
        txtPUn.Enabled = False
    Else
        txtURe.Enabled = True
        txtPUn.Enabled = True
    End If
    Calculo_Automatico
End Sub

Private Sub txtCan_GotFocus()
    Enfoque txtCan
End Sub

Private Sub txtCan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtCan) > 0 Then
            txtURe.Enabled = True
            txtURe.SetFocus
        Else
            txtCan.SetFocus
        End If
    End If
    Reales_Positivos KeyAscii, txtCan
End Sub
Private Sub Reales_Positivos(k As Integer, t As TextBox)
Dim t1 As String
    k = Asc(UCase(Chr(k)))
    If k = 8 Then Exit Sub
    If k <> 45 And k <> 44 And k <> 32 And k <> 69 And k <> 43 Then
        t1 = Left(t, t.SelStart)
        t1 = t1 & Chr(k) & Right(t, Len(t) - Len(t1))
        If IsNumeric(t1) Then Exit Sub
    End If
    k = 0
    
End Sub

Private Sub txtCan_LostFocus()
    txtCan = Format(Val(txtCan), "0.00")
End Sub

Private Sub txtordfab_GotFocus()
    Enfoque txtordfab
End Sub

Private Sub txtordfab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtCo1.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtco1_GotFocus()
    Enfoque txtCo1
End Sub

Private Sub txtCo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdOK.SetFocus
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub Enfoque(OBJ As Object)
  OBJ.SelStart = 0
  OBJ.SelLength = Len(OBJ)
End Sub


Public Function Existe(Tipo As Integer, Cod As String, Tabla As String, Campo As String, Fecha As Boolean, Optional Cod2 As String, Optional cCampo2 As String, Optional Cod3 As String, Optional cCampo3 As String, Optional Cod4 As Boolean, Optional cCampo4 As String, Optional Cod5 As String, Optional cCampo5 As String) As Boolean
Dim cSel1 As ADODB.Recordset, cSL As String
Set cSel1 = New ADODB.Recordset

 If Fecha Then
        cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
 Else
       If UCase(Tabla) = "PUNTO_VENTA" Then
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       Else
                cSL = "Select * from " & Tabla & "  Where " & Campo & " =  '" & Cod & "'"
       End If
       If Trim(Cod2) <> "" Then
            cSL = cSL & " And  " & cCampo2 & " =  '" & SupCadSQL(Cod2) & "'"
       End If
       If Trim(Cod3) <> "" Then
            cSL = cSL & " And  " & cCampo3 & " =  '" & SupCadSQL(Cod3) & "'"
       End If
       If Trim(cCampo4) <> "" Then
            If Cod4 = True Then
                cSL = cSL & " And  " & cCampo4
            Else
                cSL = cSL & " And  " & Not cCampo4
            End If
        End If
        If Trim(Cod5) <> "" Then
            cSL = cSL & " And  " & cCampo5 & " =  '" & Cod5 & "'"
        End If
 End If
 
Select Case Tipo
Case 1 'Bd. Comun
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 2 'Bd. Config
            cSel1.Open cSL, VGcnx, adOpenStatic
Case 3 'Bd. Contab
            cSel1.Open cSL, VGcnxCT, adOpenStatic
End Select

If cSel1.RecordCount > 0 Then
     Existe = True
Else
     Existe = False
End If
'csel1.Close
End Function

Public Function SupCadSQL(S As String) As String
 Dim Aux As String
 If Not IsNull(S) Then
     Aux = Replace(S, "'", "''")
 End If
 SupCadSQL = Aux
 
End Function

Private Sub txtPDe_Change()
    Calculo_Automatico
End Sub

Private Sub txtPDe_GotFocus()
    Enfoque txtPDe
End Sub

Private Sub txtPDe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtPIg.SetFocus
    End If
    Reales_Positivos KeyAscii, txtPDe
End Sub

Private Sub txtPDe_LostFocus()
    txtPDe = Format(Val(txtPDe), "0.00")
End Sub

Private Sub txtPIg_Change()
    Calculo_Automatico
End Sub

Private Sub txtPIg_GotFocus()
    Enfoque txtPIg
End Sub

Private Sub txtPIg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        txtordfab.SetFocus
    End If
    Reales_Positivos KeyAscii, txtPIg
End Sub

Private Sub txtPIg_LostFocus()
    txtPIg = Format(Val(txtPIg), "0.00")
End Sub

Private Sub txtPUn_Change()
    If Val(txtPUn) = 0 Then
        txtPDe = "0.00"
        txtPDe.Enabled = False
        txtPIg = Format(Igv, "0.00")
        txtPIg.Enabled = False
    Else
        txtPDe.Enabled = True
        txtPIg.Enabled = True
    End If
    Calculo_Automatico
End Sub

Private Sub txtPUn_GotFocus()
    Enfoque txtPUn
End Sub

Private Sub txtPUn_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtPUn) > 0 Then
            txtPDe.SetFocus
        Else
            txtPUn.SetFocus
        End If
    End If
    Reales_Positivos KeyAscii, txtPUn
End Sub

Private Sub txtPUn_LostFocus()
    txtPUn = Format(Val(txtPUn), "0.00")
End Sub

Private Sub txtRef_Change()
    If Val(txtRef) = 0 Then
        If Me.ActiveControl.Name <> "txtURe" Then
            txtPUn = "0.00"
            txtPUn.Enabled = False
        End If
    Else
        txtPUn.Enabled = True
    End If
End Sub

Private Sub txtRef_GotFocus()
    Enfoque txtRef
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(txtRef) > 0 Then
            txtPUn.SetFocus
        Else
            txtRef.SetFocus
        End If
    End If
    Reales_Positivos KeyAscii, txtRef
    Calculo_Automatico
End Sub

Private Sub txtRef_LostFocus()
    txtRef = Format(Val(txtRef), "0.00")
End Sub

Private Sub txtURe_Change()
    If txtURe <> "" Then
        txtPUn = "0.00"
        txtPUn.Enabled = False
    Else
        txtPUn.Enabled = True
    End If
    txtRef = ""
    txtRef.Enabled = False
    Calculo_Automatico
End Sub

Private Sub txtURe_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT um_abrev,um_nombre FROM tabunimed"
    Adodc2.Open strsql, VGcnx, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Unidades de Medida"
    frmReferencia.Inicio
    frmReferencia.Show vbModal
    Adodc2.Close
    
    If vGUtil(1) <> "" Then
        txtURe = vGUtil(1)
        txtURe_KeyPress 13
    End If
End Sub

Private Sub txtURe_GotFocus()
    Enfoque txtURe
End Sub

Private Sub txtURe_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtURe_DblClick
End Sub

Private Sub txtURe_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtURe = Trim(txtURe)
        If txtURe <> "" Then
            If Not Existe(1, txtURe, "tabunimed", "um_abrev", False) Then
                Mensaje = "La Unidad de medida de Referencia no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtURe.SetFocus
            Else
                If Not txtRef.Enabled Then
                    txtRef = "0.00"
                    txtRef.Enabled = True
                End If
                txtRef.SetFocus
            End If
        Else
            txtPUn.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Sub Calculo_Automatico()
    If Not activado Then Exit Sub
    
    If Not txtRef.Enabled Then
        lblDes = Format(Val(txtPUn) * Val(txtPDe) / 100, "0.00")
        lblPNe = Format(Val(txtPUn) - Val(lblDes), "0.00")
        lblTCo = Format(Val(txtCan) * Val(lblPNe), "0.00")
    Else
        lblDes = Format(Val(txtPUn) * Val(txtPDe) / 100, "0.00")
        lblPNe = Format(Val(txtPUn) - Val(lblDes), "0.00")
        lblTCo = Format(Val(txtRef) * Val(lblPNe), "0.00")
    End If
    lblIgv = Format(Val(lblTCo) * Val(txtPIg) / 100, "0.00")
    lblTNe = Format(Val(lblTCo) + Val(lblIgv), "0.00")
End Sub
