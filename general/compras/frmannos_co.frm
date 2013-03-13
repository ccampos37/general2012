VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form frmannos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aperturar Año"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   300
      Left            =   1320
      TabIndex        =   11
      Top             =   5280
      Width           =   1155
   End
   Begin RichTextLib.RichTextBox RichCom 
      Height          =   210
      Left            =   60
      TabIndex        =   10
      Top             =   5205
      Visible         =   0   'False
      Width           =   525
      _ExtentX        =   926
      _ExtentY        =   370
      _Version        =   393217
      TextRTF         =   $"frmannos_co.frx":0000
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmannos_co.frx":007A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   300
      Left            =   2580
      TabIndex        =   9
      Top             =   5280
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Height          =   885
      Left            =   60
      TabIndex        =   4
      Top             =   4305
      Width           =   4785
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   315
         Left            =   3405
         TabIndex        =   8
         Top             =   405
         Width           =   1110
      End
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "&Generar"
         Height          =   315
         Left            =   2220
         TabIndex        =   7
         Top             =   405
         Width           =   1110
      End
      Begin MSComCtl2.DTPicker DTPanno 
         Height          =   315
         Left            =   90
         TabIndex        =   5
         Top             =   405
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "Elije el Año : yyyy"
         Format          =   23920643
         UpDown          =   -1  'True
         CurrentDate     =   37491
      End
      Begin VB.Label Label3 
         Caption         =   "Año"
         Height          =   165
         Left            =   105
         TabIndex        =   6
         Top             =   195
         Width           =   1815
      End
   End
   Begin MSComctlLib.ListView LV_Meses 
      Height          =   3975
      Left            =   1995
      TabIndex        =   1
      Top             =   285
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   7011
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Meses"
         Object.Width           =   4304
      EndProperty
   End
   Begin TrueOleDBGrid70.TDBGrid TDB_Anno 
      Height          =   3930
      Left            =   45
      TabIndex        =   0
      Top             =   315
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   6932
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Años"
      Columns(0).DataField=   "generalanno"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   1
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).ScrollBars=   2
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=1"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2064"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1984"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      Appearance      =   2
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
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
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=1,.bold=0,.fontsize=825,.italic=0"
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
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=32,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=29,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=30,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=31,.parent=17"
      _StyleDefs(40)  =   "Named:id=33:Normal"
      _StyleDefs(41)  =   ":id=33,.parent=0"
      _StyleDefs(42)  =   "Named:id=34:Heading"
      _StyleDefs(43)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(44)  =   ":id=34,.wraptext=-1"
      _StyleDefs(45)  =   "Named:id=35:Footing"
      _StyleDefs(46)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(47)  =   "Named:id=36:Selected"
      _StyleDefs(48)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(49)  =   "Named:id=37:Caption"
      _StyleDefs(50)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(51)  =   "Named:id=38:HighlightRow"
      _StyleDefs(52)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=39:EvenRow"
      _StyleDefs(54)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(55)  =   "Named:id=40:OddRow"
      _StyleDefs(56)  =   ":id=40,.parent=33"
      _StyleDefs(57)  =   "Named:id=41:RecordSelector"
      _StyleDefs(58)  =   ":id=41,.parent=34"
      _StyleDefs(59)  =   "Named:id=42:FilterBar"
      _StyleDefs(60)  =   ":id=42,.parent=33"
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Lista de meses"
      Height          =   255
      Left            =   2025
      TabIndex        =   3
      Top             =   15
      Width           =   2835
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Años aperturados"
      Height          =   255
      Left            =   45
      TabIndex        =   2
      Top             =   15
      Width           =   1935
   End
End
Attribute VB_Name = "frmannos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsannos As ADODB.Recordset
Dim SQL As String
 
Private Sub CmdAceptar_Click()
On Error GoTo Actmes
If rsannos.RecordCount = 0 Then Exit Sub
VGCNx.BeginTrans
SQL = "UPDATE dbo.CT_General set " & ActualizaMeses & " Where " & _
      "generalanno='" & rsannos!generalanno & "'"
VGCNx.Execute SQL
VGCNx.CommitTrans
Unload Me
Exit Sub
Actmes:
    VGCNx.RollbackTrans
    MsgBox "Hubo Errores al Actualizar los meses " & Chr(13) & _
           Err.Description
        
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub
Private Function ActualizaMeses() As String
Dim I As Integer
    SQL = ""
    For I = 1 To 12
        SQL = SQL + "generalmes" & Format(I, "00") & "=" & IIf(LV_Meses.ListItems(I).Checked, 1, 0)
        If I <= 11 Then SQL = SQL + ","
    Next
    ActualizaMeses = SQL
End Function
Public Sub cmdGenerar_Click()
'Cargando la plantilla de las tablasç
On Error GoTo xtrans
    If Not VeriExisteAño Then Exit Sub
    
    RichCom.FileName = App.Path & "\plantillatablasanual.sql"
    If Trim(RichCom.FileName) = "" Then
        MsgBox "No se ha encontrado el archivo " & App.Path & "\" & " plantillatablasano.sql " & Chr(13) & _
               "en la Ruta especificada", vbExclamation
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    RichCom.Text = Replace(RichCom.Text, "XXXX", Format(Year(DTPanno.Value), "0000"))
    VGCNx.BeginTrans
    Call EjecutarLote(RichCom, VGCNx)
    'Grabando el Año
    SQL = "Insert Into co_correlames(ano,mes01,mes02,mes03,mes04,mes05,mes06,mes07,mes08,mes09,mes10,mes11,mes12)"
    SQL = SQL & " values('" & Year(DTPanno.Value) & "',0,0,0,0,0,0,0,0,0,0,0,0)"
    VGCNx.Execute SQL
    VGCNx.CommitTrans
    'Cargando los Meses por primera vez
    If rsannos.RecordCount = 0 Then
        Call CargarMeses
    End If
    rsannos.Requery
    Screen.MousePointer = vbDefault
    MsgBox "El año se genero Satisfactoriamente ", vbInformation
    Exit Sub
xtrans:
    Screen.MousePointer = vbDefault
    VGCNx.RollbackTrans
    MsgBox "Hubo Errores al generar el nuevo año " & Chr(13) & _
           Err.Description
End Sub
Private Function VeriExisteAño() As Boolean
Dim rsaux As ADODB.Recordset
    Set rsaux = New ADODB.Recordset
    VeriExisteAño = False
    rsaux.Open "select * from ct_general where generalanno='" & Year(DTPanno.Value) & "'", VGCNx, adOpenKeyset, adLockReadOnly
    If rsaux.RecordCount > 0 Then
        MsgBox "El año de ejercicio " & Year(DTPanno.Value) & " ya se genero ", vbInformation
        Exit Function
    End If
    VeriExisteAño = True
End Function

Private Sub Form_Load()
    Me.Height = 6060
    Me.Width = 4980
    Set rsannos = New ADODB.Recordset
    rsannos.Open "ct_general", VGCNx, adOpenKeyset, adLockOptimistic
    If rsannos.RecordCount > 0 Then
        Call CargarMeses
    End If
    Set TDB_Anno.DataSource = rsannos
End Sub
Private Sub CargarMeses()
Dim items As ListItems, I As Integer
    Set VGvardllgen = New dllgeneral.dll_general
    Set items = LV_Meses.ListItems
    For I = 1 To 12
        items.Add I, Format(I, "C0"), VGvardllgen.DESMES(Format(I, "00")), , 1
    Next
End Sub
Private Sub ActivarMeses()
Dim I As Integer
    For I = 1 To rsannos.Fields.Count - 1
        LV_Meses.ListItems(I).Checked = rsannos.Fields(I).Value
    Next
End Sub

Private Sub LV_Meses_ItemCheck(ByVal Item As MSComctlLib.ListItem)
   'Item.Checked = Not Item.Checked
End Sub

Private Sub TDB_Anno_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
    If rsannos.RecordCount = 0 Then Exit Sub
    Call ActivarMeses
End Sub
