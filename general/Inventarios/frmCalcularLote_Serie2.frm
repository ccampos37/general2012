VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalcularLote_Serie2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calcular Lotes"
   ClientHeight    =   2025
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6855
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      Begin VB.ComboBox Combo1 
         Height          =   288
         ItemData        =   "frmCalcularLote_Serie2.frx":0000
         Left            =   1080
         List            =   "frmCalcularLote_Serie2.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   270
         Width           =   4335
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   150
         TabIndex        =   3
         Top             =   1110
         Width           =   5355
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   315
            Left            =   90
            TabIndex        =   4
            Top             =   210
            Width           =   5145
            _ExtentX        =   9075
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Total ==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1080
         TabIndex        =   8
         Top             =   780
         Width           =   1395
      End
      Begin VB.Label Label2 
         Caption         =   "Transcurridos==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2580
         TabIndex        =   7
         Top             =   810
         Width           =   1635
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4110
         TabIndex        =   6
         Top             =   810
         Width           =   1365
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   645
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1104
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   645
      Left            =   5808
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   144
      Width           =   795
   End
End
Attribute VB_Name = "frmCalcularLote_Serie2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit
Dim PCount As Long
Dim cConexAux As ADODB.Connection
Dim adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim rs As ADODB.Recordset
Dim cRt As String
Dim almacen As String
Dim nTra As Integer

Private Sub cmd_exit_Click()
Unload Me
End Sub

Private Sub Combo2_Change()

End Sub

Private Sub Command1_Click()
  Dim rs1 As New ADODB.Recordset
  Dim r1, r2, r3 As Double
  Dim ncodi, nalma, nlote, nsql As String
  Dim n As Double
  Dim nflag As Integer
  Dim Text2 As String * 2
  
  Text2 = "" & Combo1.text
  nflag = 0

  Set rs = Nothing
  

  Set rs1 = VGCNx.Execute("select * from sysobjects where name like 'xx_tempora%'")
  If rs1.RecordCount > 0 Then
     VGCNx.Execute "drop table [dbo].[xx_tempora]"
  End If
  rs1.Close
  Set rs1 = Nothing
  
  Set rs1 = VGCNx.Execute("select * from sysobjects where name like 'xx_i0%'")
  If rs1.RecordCount > 0 Then
     VGCNx.Execute "drop table [dbo].[xx_i0]"
  End If
  rs1.Close
  Set rs1 = Nothing

   VGCNx.Execute "SELECT STSALMA,STSCODIGO,STSLOTE " & _
              " INTO xx_TEMPORA FROM STKLOTE " & _
              " inner join movalmdet " & _
              " on dealma=stsalma and decodigo=stscodigo and delote=stslote " & _
              " where stsalma='" & Trim(Text2) & "' " & _
              " group by stsalma,stscodigo,stslote "

  
  nsql = "select decodigo,delote,round(sum(case catipmov when 'I' then round(decantid,2) else 0 end),2) as ingreso," & _
         " round(sum(case catipmov when 'S' then round(decantid,2) else 0 end),2) as salida " & _
         " into xx_i0 " & _
         " from movalmdet inner join movalmcab " & _
         " on dealma=caalma and detd=catd and denumdoc=canumdoc " & _
         " where dealma='" & Trim(Text2) & "'" & _
         " group by decodigo,delote " & _
         "  order by decodigo,delote"
         
         
  VGCNx.Execute nsql
  
  Bar1.Value = 0
  Label3 = Format(0, "###,##0.00")
  Label4 = Format(0, "###,##0.00")
  DoEvents
  n = 0
  rs1.Open "xx_tempora", VGCNx, adOpenDynamic, adLockOptimistic
  If rs1.RecordCount > 0 Then
    rs1.MoveLast
    Bar1.Max = rs1.RecordCount
    Label3 = Format(rs1.RecordCount, "###,##0.00")
    VGCNx.BeginTrans
    nflag = 1
    rs1.MoveFirst
    Do Until rs1.EOF
        nalma = rs1!stsalma
        ncodi = rs1!stscodigo
        nlote = rs1!stslote
        Set rs = VGCNx.Execute("select * from xx_i0 where decodigo='" & ncodi & "' and delote='" & nlote & "'")
        r1 = 0: r2 = 0: r3 = 0
        If rs.RecordCount > 0 Then
           r1 = rs!ingreso
           r2 = rs!Salida
        End If
        rs.Close
        Set rs = Nothing
        r3 = r1 - r2
        VGCNx.Execute "update stklote " & _
                   " set stslkdis=" & r3 & _
                   " where stsalma='" & nalma & "' and stscodigo='" & ncodi & "' and stslote='" & nlote & "'"
        n = n + 1
        Label4 = Format(n, "###,##0.00")
        Bar1.Value = n
        DoEvents
        rs1.MoveNext
    Loop
    VGCNx.CommitTrans
    nflag = 0
  End If
  rs1.Close
  Set rs1 = Nothing
  
 ' Set cn = Nothing
  MsgBox "Proceso Terminado Satisfactoriamente..!!", vbInformation, "AVISO"
  Unload Me
   
nerror:
    If Err Then
        If nflag = 1 Then
            VGCNx.RollbackTrans
        End If
        MsgBox "Error : " & Err.Number & "-" & Err.Description
        Err = 0
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
 central Me
 Call Carga_Almacen
End Sub

Private Sub Carga_Almacen()
  Dim rsc As New ADODB.Recordset
  Combo1.Clear
  Set rsc = VGCNx.Execute("select TAALMA,TADESCRI from tabalm ")
  If rsc.RecordCount > 0 Then
      rsc.MoveFirst
      Do Until rsc.EOF
        Combo1.AddItem rsc.Fields(0) & "-" & IIf(IsNull(rsc.Fields(1)), "****", rsc.Fields(1))
        rsc.MoveNext
      Loop
  End If
  rsc.Close
  Set rsc = Nothing
End Sub

Sub Cmd_RestoreSaldos_Click()
  Dim rs1 As New ADODB.Recordset
  Dim r1, r2, r3 As Double
  Dim ncodi, nalma, nlote, nsql As String
  Dim n As Double
  Dim nflag As Integer
  Dim Text2 As String * 2
  
  Text2 = "" & Combo1.text
  nflag = 0

  '     Set VGcnx  = Nothing
  Set rs = Nothing
  
 
  Set rs1 = VGCNx.Execute("select * from sysobjects where name like 'xx_tempora%'")
  If rs1.RecordCount > 0 Then
     VGCNx.Execute "drop table [dbo].[xx_tempora]"
  End If
  rs1.Close
  Set rs1 = Nothing
  
  Set rs1 = VGCNx.Execute("select * from sysobjects where name like 'xx_i0%'")
  If rs1.RecordCount > 0 Then
     VGCNx.Execute "drop table [dbo].[xx_i0]"
  End If
  rs1.Close
  Set rs1 = Nothing

  nsql = "SELECT STSALMA,STSCODIGO,STSLOTE " & _
              " INTO xx_TEMPORA FROM STKLOTE " & _
              " inner join movalmdet " & _
              " on dealma=stsalma and decodigo=stscodigo and delote=stslote " & _
              " where stsalma='" & Trim(Text2) & "' " & _
              " group by stsalma,stscodigo,stslote "

  VGCNx.Execute nsql
  
  nsql = "select decodigo,delote,round(sum(case catipmov when 'I' then round(decantid,2) else 0 end),2) as ingreso," & _
         " round(sum(case catipmov when 'S' then round(decantid,2) else 0 end),2) as salida " & _
         " into xx_i0 " & _
         " from movalmdet inner join movalmcab " & _
         " on dealma=caalma and detd=catd and denumdoc=canumdoc " & _
         " where dealma='" & Trim(Text2) & "'" & _
         " group by decodigo,delote " & _
         "  order by decodigo,delote"
         
         
  VGCNx.Execute nsql
  
  Bar1.Value = 0
  'Label3 = Format(0, "###,##0.00")
  'Label4 = Format(0, "###,##0.00")
  DoEvents
  n = 0
  rs1.Open "xx_tempora", VGCNx, adOpenDynamic, adLockOptimistic
  If rs1.RecordCount > 0 Then
    rs1.MoveLast
    Bar1.Max = rs1.RecordCount
    'Label3 = Format(rs.RecordCount, "###,##0.00")
    VGCNx.BeginTrans
    nflag = 1
    rs1.MoveFirst
    Do Until rs1.EOF
        nalma = rs1!stsalma
        ncodi = rs1!stscodigo
        nlote = rs1!stslote
        Set rs = VGCNx.Execute("select * from xx_i0 where decodigo='" & ncodi & "' and decodigo='" & nlote & "'")
        r1 = 0: r2 = 0: r3 = 0
        If rs.RecordCount > 0 Then
           r1 = rs!ingreso
           r2 = rs!Salida
        End If
        rs.Close
        Set rs = Nothing
        r3 = r1 - r2
        VGCNx.Execute "update stklote " & _
                   " set stslkdis=" & r3 & _
                   " where stsalma='" & nalma & "' and stscodigo='" & ncodi & "' and stslote='" & nlote & "'"
        n = n + 1
        'Label4 = Format(n, "###,##0.00")
        'Bar1.Value = n
        DoEvents
        rs1.MoveNext
    Loop
    VGCNx.CommitTrans
    nflag = 0
  End If
  rs1.Close
  Set rs1 = Nothing

  MsgBox "Proceso Terminado Satisfactoriamente..!!", vbInformation, "AVISO"
  Unload Me
  
nerror:
    If Err Then
        If nflag = 1 Then
            VGCNx.RollbackTrans
        End If
        MsgBox "Error : " & Err.Number & "-" & Err.Description
        Err = 0
        Exit Sub
        
    End If

    
End Sub

