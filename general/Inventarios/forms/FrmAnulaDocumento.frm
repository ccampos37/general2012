VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDatGrd.ocx"
Begin VB.Form FrmAnulaDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Anulacion de Documentos"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9780
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   9780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Height          =   2595
      Left            =   240
      TabIndex        =   17
      Top             =   3480
      Width           =   9195
      Begin MSDataGridLib.DataGrid DGrid2 
         Height          =   1575
         Left            =   240
         TabIndex        =   24
         Top             =   840
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2778
         _Version        =   393216
         BackColor       =   12648384
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin VB.Label Label2 
         Caption         =   "Tipo"
         Height          =   195
         Index           =   7
         Left            =   6360
         TabIndex        =   29
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   6780
         TabIndex        =   28
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Nro Documento"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   23
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   1440
         TabIndex        =   22
         Top             =   270
         Width           =   1785
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Documento"
         Height          =   195
         Index           =   4
         Left            =   3300
         TabIndex        =   21
         Top             =   330
         Width           =   1425
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   4770
         TabIndex        =   20
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   195
         Index           =   3
         Left            =   7140
         TabIndex        =   19
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   8280
         TabIndex        =   18
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2595
      Left            =   210
      TabIndex        =   9
      Top             =   840
      Width           =   9195
      Begin MSDataGridLib.DataGrid DGrid1 
         Height          =   1575
         Left            =   120
         TabIndex        =   25
         Top             =   720
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   2778
         _Version        =   393216
         BackColor       =   12648384
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
               LCID            =   3082
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
               LCID            =   3082
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
      Begin VB.Label Label2 
         Caption         =   "Tipo"
         Height          =   195
         Index           =   6
         Left            =   6120
         TabIndex        =   27
         Top             =   300
         Width           =   315
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FF0000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   6540
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   8280
         TabIndex        =   16
         Top             =   270
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Transaccion"
         Height          =   195
         Index           =   2
         Left            =   7140
         TabIndex        =   15
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   4530
         TabIndex        =   14
         Top             =   270
         Width           =   1515
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Documento"
         Height          =   195
         Index           =   1
         Left            =   3060
         TabIndex        =   13
         Top             =   330
         Width           =   1425
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00C0FFFF&
         Height          =   345
         Left            =   840
         TabIndex        =   12
         Top             =   270
         Width           =   2025
      End
      Begin VB.Label Label2 
         Caption         =   "No Guia "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   330
         Width           =   585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   180
      TabIndex        =   2
      Top             =   30
      Width           =   9225
      Begin VB.CommandButton Command3 
         Caption         =   "&Consulta"
         Height          =   345
         Left            =   8250
         TabIndex        =   10
         Top             =   330
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   870
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   3105
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   6840
         MaxLength       =   11
         TabIndex        =   6
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   6360
         MaxLength       =   2
         TabIndex        =   5
         Top             =   330
         Width           =   375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   5100
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Documento"
         Height          =   195
         Index           =   0
         Left            =   4080
         TabIndex        =   3
         Top             =   390
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   645
      Left            =   5235
      Picture         =   "FrmAnulaDocumento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6240
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   645
      Left            =   3960
      Picture         =   "FrmAnulaDocumento.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6240
      Width           =   795
   End
End
Attribute VB_Name = "FrmAnulaDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Dim rs2 As New ADODB.Recordset
Dim nrotransf As String
Dim ok As Integer

Private Sub Command1_Click()
    Dim ncanti As Double
    Dim ncodi As String
    Dim nalma As String
    Dim nOpc As String * 1
    Dim nflag As Integer

    
    Command1.Enabled = False
    nflag = 0
    If rs1.RecordCount > 0 Then
       If MsgBox("Desea Anular el Registro?", vbYesNo, "AVISO") = vbYes Then
          Call anula(rs1, Label9)
          If VGtransf = 1 Then Call anula(rs2, Label10)
          nflag = 0
          MsgBox "El documento ha sido anulado satisfactoriamente...!!!", vbInformation, "AVISO"
       End If
   End If
End Sub
Private Sub anula(rs As Recordset, tipo As String)
 On Error GoTo nerror
 nalma = Left(Combo2, 2)
 rs.MoveFirst
 VGCNx.BeginTrans
 nflag = 1
 Do Until rs.EOF
    ncodi = "" & Trim(rs!acodigo)
    ncanti = IIf(IsNull(rs!DECANTID), 0, rs!DECANTID)
    If tipo = "I" Then
       VGCNx.Execute "UPDATE stkart " & _
                     " set stskdis=stskdis-" & ncanti & _
                     " Where stalma='" & rs!dealma & "' and stcodigo='" & ncodi & "'"
     ElseIf tipo = "S" Then
            VGCNx.Execute "UPDATE stkart " & _
                          " set stskdis=stskdis+" & ncanti & _
                          " Where stalma='" & rs!dealma & "' and stcodigo='" & ncodi & "'"
     End If
    rs.MoveNext
Loop
rs.MoveFirst
SQL = "UPDATE movalmcab SET  casitgui='A', usuariomodifica='" & UCase(VGUsuario) & "'"
SQL = SQL & " Where caalma='" & rs!dealma & "' and catd='" & rs!detd & "'"
SQL = SQL & " and canumdoc='" & rs!denumdoc & "'"
VGCNx.Execute (SQL)
 '            VGCNx.Execute "UPDATE movalmcab " & _
 '                             " SET  casitgui='A', usuariomodifica='" & UCase(VGUsuario) & "'" & _
 '                             " Where CAtd='" & Trim(Text1) & "' and CAnumdoc='" & Trim(Text2) & "' and CATIPMOV='" & Left(Combo1, 1) & "' and caalma='" & Left(Combo2, 2) & "' "
            
VGCNx.CommitTrans
Exit Sub
nerror:
    If Err Then
        If nflag = 1 Then
            VGCNx.RollbackTrans
        End If
        MsgBox "Error : " & Err.Number & "-" & Err.Description
        Err = 0
        Exit Sub
        Resume
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
      ok = 0
   If VGtransf = 0 Then
      Call selecciona
    Else
      Call selecciona
      If ok = 1 Then Call selecciona2
   End If
End Sub
Private Sub selecciona()
   Dim ndato As String * 1
   Set rs = Nothing
   Set DGrid1.DataSource = Nothing
   DGrid1.ClearFields
   
   Label3 = ""
   Label4 = ""
   Label5 = ""
   
   Set rs = VGCNx.Execute("select * from movalmcab where CAtd='" & Trim(Text1) & "' and CAnumdoc='" & Trim(Text2) & "' and CATIPMOV='" & Left(Combo1, 1) & "' and caalma='" & Left(Combo2, 2) & "'")
   If rs.RecordCount > 0 Then
       ndato = "" & rs!Casitgui
       
       If ndato = "A" Then
            MsgBox "El documento esta anulado...Verifique!!", vbInformation, "AVISO"
            Exit Sub
       End If
       nrotransf = ESNULO(rs!caNROtransf, "")
       If nrotransf = "" And VGtransf = 1 Then
            MsgBox "El documento no es de Transferencia ...Verifique!!", vbInformation, "AVISO"
            Exit Sub
       End If
       Command1.Enabled = True
       Label3 = rs!CARFTDOC & "-" & rs!CARFNDOC
       
       Label4 = "" & Format(rs!CAFECDOC, "dd/mm/yyyy")
       Label5 = "" & rs!cacodmov
       Label9 = rs!catipmov
       ok = 1
       Set rs1 = VGCNx.Execute("select acodigo ,adescri,dealma,decantid,decanref1,detd,denumdoc from movalmcab A inner join movalmdet B" & _
                             " on a.caalma=b.dealma and a.catd=b.detd and a.canumdoc=b.denumdoc " & _
                             " inner join maeart c " & _
                             " on b.decodigo=c.acodigo " & _
                             " where A.CATD='" & Trim(Text1) & "' and A.CANUMDOC='" & Trim(Text2) & "' and A.CATIPMOV='" & Left(Combo1, 1) & "' and A.caalma='" & Left(Combo2, 2) & "'")
        If rs1.RecordCount > 0 Then

        Call cargar
        End If
        'rs.Close
        
    Else
        MsgBox "No existe Documento � esta anulado...Verifique!!!", vbInformation, "AVISO"
    End If
    rs.Close
    Set rs = Nothing
  End Sub
Private Sub selecciona2()
      Dim ndato As String * 1
   
   Set rs = Nothing
   Set DGrid2.DataSource = Nothing
   DGrid2.ClearFields
   
   Label6 = ""
   Label7 = ""
   Label8 = ""
   
   Set rs = VGCNx.Execute("select * from movalmcab where CAtipotransf='TR' and CAnrotransf='" & nrotransf & "' and CATIPMOV='I' ")
   If rs.RecordCount > 0 Then
       ndato = "" & rs!Casitgui
       If ndato = "A" Then
            MsgBox "El documento esta anulado...Verifique!!", vbInformation, "AVISO"
            Exit Sub
       End If
       Command1.Enabled = True
       Label8 = rs!CATD & "-" & rs!CAnumDOC
       Label7 = "" & Format(rs!CAFECDOC, "dd/mm/yyyy")
       Label6 = "" & rs!cacodmov
       Label10 = rs!catipmov
       Set rs2 = VGCNx.Execute("select acodigo ,adescri,dealma,decantid,decanref1,detd,denumdoc from movalmcab A inner join movalmdet B" & _
                             " on a.caalma=b.dealma and a.catd=b.detd and a.canumdoc=b.denumdoc " & _
                             " inner join maeart c " & _
                             " on b.decodigo=c.acodigo " & _
                             " where A.CAtipotransf='TR' and A.canrotransf='" & nrotransf & "' and A.catipmov='I'")
        If rs2.RecordCount > 0 Then
           Call cargar2
        End If
        'rs.Close
        
    Else
        MsgBox "No existe Documento � esta anulado...Verifique!!!", vbInformation, "AVISO"
    End If
    rs.Close
    Set rs = Nothing
  End Sub
  

Private Sub Form_Load()
  Dim rsc As New ADODB.Recordset
  SQL = "select TAALMA,TADESCRI,'','' from tabalm where taalma='*'"
  Set rs = VGCNx.Execute(SQL)
  
  Combo2.Clear
  If VGParametros.listaPuntoVtas <> "" Then
     SQL = "select TAALMA,TADESCRI from tabalm where puntovtacodigo in (" & VGParametros.listaPuntoVtas & ")"
    Else
     SQL = "select TAALMA,TADESCRI from tabalm "
  End If
  Set rsc = VGCNx.Execute(SQL)
  If rsc.RecordCount > 0 Then
      rsc.MoveFirst
      Do Until rsc.EOF
        Combo2.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
        rsc.MoveNext
      Loop
      Combo2.ListIndex = 0
  End If
  rsc.Close
  Set rsc = Nothing
  If VGtransf = 0 Then
     Combo1.Clear
     Combo1.AddItem "I-Ingreso"
     Combo1.AddItem "S-Salida"
   Else
     Combo1.Clear
     Combo1.AddItem "S-Salida"
  End If
  Command1.Enabled = False

End Sub

Sub cargar()

        Set DGrid1.DataSource = rs1
        With DGrid1
            .Columns(0).Caption = "Codigo"
            .Columns(0).Width = 1200
            .Columns(1).Caption = "Descripcion"
            .Columns(1).Width = 4800
            .Columns(2).Caption = "Cantidad"
            .Columns(2).NumberFormat = "##,###,#0.00"
            .Columns(2).Alignment = dbgRight
            .Columns(2).Width = 1200
            .Columns(3).Caption = "Can.Ref"
            .Columns(3).NumberFormat = "##,###,#0.00"
            .Columns(3).Alignment = dbgRight
            .Columns(3).Width = 1200
            .Refresh
        End With
End Sub

Sub cargar2()
        Set DGrid2.DataSource = rs2
        With DGrid2
            .Columns(0).Caption = "Codigo"
            .Columns(0).Width = 1200
            .Columns(1).Caption = "Descripcion"
            .Columns(1).Width = 4800
            .Columns(2).Caption = "Cantidad"
            .Columns(2).NumberFormat = "##,###,#0.00"
            .Columns(2).Alignment = dbgRight
            .Columns(2).Width = 1200
            .Columns(3).Caption = "Can.Ref"
            .Columns(3).NumberFormat = "##,###,#0.00"
            .Columns(3).Alignment = dbgRight
            .Columns(3).Width = 1200
            .Refresh
        End With

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    Text1 = UCase(Text1)
    Text2.SetFocus
  End If
End Sub

Private Sub Text1_LostFocus()
    Call Text1_KeyPress(13)
End Sub

Private Sub Text2_LostFocus()
If IsNumeric(Text2.text) Then
   Text2.text = Right("00000000000" + RTrim(Text2.text), Text2.MaxLength)
   Command3.SetFocus
End If
End Sub
