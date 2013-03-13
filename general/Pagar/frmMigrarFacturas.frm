VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmMigrarFacturas 
   Caption         =   "Migrar Facturas desde Contabilidad"
   ClientHeight    =   2565
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2565
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin TrueOleDBGrid70.TDBGrid TDBGrid1 
      Height          =   2025
      Left            =   90
      TabIndex        =   4
      Top             =   2670
      Width           =   6390
      _ExtentX        =   11271
      _ExtentY        =   3572
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
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
   Begin VB.TextBox txtMes 
      Height          =   285
      Left            =   1530
      TabIndex        =   3
      Top             =   180
      Width           =   705
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   240
      Left            =   120
      ScaleHeight     =   180
      ScaleWidth      =   4335
      TabIndex        =   1
      Top             =   2160
      Width           =   4395
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar Migracion"
      Height          =   420
      Left            =   1290
      TabIndex        =   0
      Top             =   1005
      Width           =   1995
   End
   Begin VB.Label Label3 
      Caption         =   "0"
      Height          =   240
      Left            =   2085
      TabIndex        =   6
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label Label2 
      Caption         =   "Total Facturas Procesar"
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1665
      Width           =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "Indicar Mes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   255
      TabIndex        =   2
      Top             =   195
      Width           =   1230
   End
End
Attribute VB_Name = "frmMigrarFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim sqlcad As String

Private Sub cmdProcesar_Click()
   Call GrabarData
End Sub

Private Sub Form_Load()
  'sqlcad = "select * from ct_detcomprob2002 where cabcomprobmes=12 and cuentacodigo like '421%' and documentocodigo<>'03' order by 2 "
  sqlcad = "select * from ct_detcomprob2002 where cabcomprobmes=12 and cuentacodigo like '421%' and documentocodigo<>'03' order by 2 "
  Set rs = New ADODB.Recordset
  Set rs = VGcnxCT.Execute(sqlcad)
  'Set TDBGrid1.DataSource = rs
  ProgressBar1.Visible = False
  
End Sub

Sub GrabarData()
'On Error GoTo X
 Dim rb As New ADODB.Recordset
 Dim acmd As New ADODB.Command
 Dim xnumplan As Long
 Dim xMonto As Double
 Dim xSerie, xNumero, xcargo, xCodCliente As String
 Dim nConta As Integer, nPos As Integer
 
      nConta = 1
      ProgressBar1.Visible = True
      
      If rs.RecordCount > 0 Then
         Label3.Caption = rs.RecordCount
         VGCNx.Execute ("Delete from cp_cargo where usuariocodigo='05'")
         
         ProgressBar1.Max = rs.RecordCount
         rs.MoveFirst
         
         Do Until rs.EOF
           VGCNx.BeginTrans
            If nConta = 1 Then
               Set rb = VGCNx.Execute("select * from cp_tipoplanilla where tplanillacodigo='02'")
               If rb.RecordCount > 0 Then
                  xnumplan = Val(Trim(rb!tplanillanumerador)) + 1
               Else
                  xnumplan = 1
               End If
               rb.Close
               Set rb = Nothing
               VGCNx.Execute "update cp_tipoplanilla " & _
                          " set tplanillanumerador='" & xnumplan & "' " & _
                          " where tplanillacodigo='02'"
            End If
           
            Set rb = VGCNx.Execute("select * from cp_tipodocumento where tdocumentocodigo='" & rs.Fields("documentocodigo") & "'")
            If rb.RecordCount > 0 Then
               xcargo = rb!tdocumentotipo
            Else
               xcargo = ""
            End If
            rb.Close
            Set rb = Nothing
              If rs("monedacodigo") = g_TipoDolar Then
                 xMonto = rs.Fields("detcomprobusshaber") + rs.Fields("detcomprobussdebe")
              Else
                 xMonto = rs.Fields("detcomprobdebe") + rs.Fields("detcomprobhaber")
              End If
              nPos = InStr(1, rs("detcomprobnumdocumento"), "-")
              If nPos > 0 Then
                 xSerie = Trim(Left(rs("detcomprobnumdocumento"), nPos - 1)):
                 xSerie = Right(Format("000" & xSerie, "000"), 3)
                 xNumero = Mid(rs("detcomprobnumdocumento"), nPos + 1, Len(rs("detcomprobnumdocumento")) - nPos)
                 xNumero = Format("00000000" & xNumero, "00000000")
              Else
                 xSerie = "000"
                 xNumero = Format("00000000" & Trim(rs("detcomprobnumdocumento")), "00000000")
              End If
              xCodCliente = Escadena(Trim(Left(rs.Fields("analiticocodigo"), 11)))
              Set acmd.ActiveConnection = VGgeneral
              acmd.CommandText = "cp_ingresavarios_pro"
              acmd.CommandType = adCmdStoredProc
              acmd.Prepared = True
              With acmd
                .Parameters("@base") = "VENTAS_PRUEBA"
                .Parameters("@tipo") = "1"
                .Parameters("@tabla") = "cp_cargo"
                .Parameters("@tipodocu") = Escadena(rs.Fields("documentocodigo"))
                .Parameters("@numero") = xSerie & Trim(Right(xNumero, 8))
                .Parameters("@cliente") = xCodCliente
                .Parameters("@vendedor") = "001"
                .Parameters("@zona") = "01"
                .Parameters("@apefecemi") = rs.Fields("detcomprobfechaemision")
                .Parameters("@moneda") = rs.Fields("monedacodigo")
                .Parameters("@apeimppag") = xMonto
                .Parameters("@usuario") = "05"
                .Parameters("@tipocambio") = rs("detcomprobtipocambio")
                .Parameters("@fechaact") = Date
                .Parameters("@flagcancel") = 0
                .Parameters("@tipoplanilla") = "02"
                .Parameters("@planilla") = Format("000000" & xnumplan, "000000")
                .Parameters("@vencimiento") = rs.Fields("detcomprobfechaemision")
                .Parameters("@fechaplani") = "01/12/2002"
                .Parameters("@banco") = ""
                .Parameters("@cargoabono") = xcargo
              End With
              acmd.Execute
              Set acmd = Nothing
              DoEvents
                                
            '**** Actualizamos Saldos del cliente
            If rs.Fields("monedacodigo") = g_TipoDolar Then
               VGCNx.Execute "Update  cp_proveedor Set clientesaldodolares=isnull(clientesaldodolares,0)+" & CDbl(xMonto) & _
                    " Where clientecodigo='" & xCodCliente & "'"
            ElseIf rs.Fields("monedacodigo") = g_TipoSol Then
               VGCNx.Execute "Update  cp_proveedor Set clientesaldosoles=isnull(clientesaldosoles,0)+" & CDbl(xMonto) & _
                    " Where clientecodigo='" & xCodCliente & "'"
            End If
            VGCNx.CommitTrans
            ProgressBar1.Value = rs.AbsolutePosition
            nConta = nConta + 1
            If nConta = 21 Then nConta = 1
                                            
            rs.MoveNext
            Loop
      End If
      ProgressBar1.Value = rs.RecordCount
      ProgressBar1.Visible = False
      rs.Close
      Set rs = Nothing
      MsgBox "Los datos han sido grabados satisfactoriamente...!!!", vbInformation, MsgTitle
      Exit Sub
X:
  MsgBox "Error al Grabar: " & Err.Description, vbExclamation, Caption
  VGCNx.RollbackTrans

End Sub
