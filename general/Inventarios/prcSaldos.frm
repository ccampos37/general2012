VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form PrcSaldos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Proceso de Regeneracion de Saldos"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   645
      Left            =   1980
      Picture         =   "prcSaldos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2130
      Width           =   795
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   645
      Left            =   3255
      Picture         =   "prcSaldos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2130
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Height          =   1875
      Left            =   210
      TabIndex        =   0
      Top             =   150
      Width           =   5655
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   150
         TabIndex        =   5
         Top             =   1110
         Width           =   5355
         Begin MSComctlLib.ProgressBar Bar1 
            Height          =   315
            Left            =   90
            TabIndex        =   6
            Top             =   210
            Width           =   5145
            _ExtentX        =   9075
            _ExtentY        =   556
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
         End
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   4335
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4110
         TabIndex        =   10
         Top             =   810
         Width           =   1365
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
         TabIndex        =   9
         Top             =   810
         Width           =   1635
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
         TabIndex        =   7
         Top             =   780
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Almacen"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "PrcSaldos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
  Dim rs1 As New ADODB.Recordset
  Dim r1, r2, r3 As Double
  Dim ncodi, nalma As String
  Dim n As Double
  Dim nflag As Integer
  Dim Text2 As String * 2
  
  Text2 = "" & Combo1.text
  nflag = 0

  Set rs1 = Nothing
  Set rs1 = VGCNx.Execute("select * from sysobjects where name like 'tempora%'")
  If rs1.RecordCount > 0 Then
     VGCNx.Execute "drop table [dbo].[tempora]"
  End If
  rs1.Close
  Set rs1 = Nothing
  
  Set rs1 = VGCNx.Execute("select * from sysobjects where name like 'i0%'")
  If rs1.RecordCount > 0 Then
     VGCNx.Execute "drop table [dbo].[i0]"
  End If
  rs1.Close
  Set rs1 = Nothing

   VGCNx.Execute "SELECT STALMA,STCODIGO " & _
              " INTO TEMPORA FROM STKART " & _
              " inner join movalmdet " & _
              " on dealma=stalma and decodigo=stcodigo " & _
              " where stalma='" & Trim(Text2) & "' " & _
              " group by stalma,stcodigo "

  VGCNx.Execute ("update stkart set stskdis=0 where stalma='" & Trim(Text2) & "'")
  
  nsql = "select decodigo,round(sum(case catipmov when 'I' then round(decantid,2) else 0 end),2) as ingreso," & _
         " round(sum(case catipmov when 'S' then round(decantid,2) else 0 end),2) as salida " & _
         " into i0 " & _
         " from movalmdet inner join movalmcab " & _
         " on dealma=caalma and detd=catd and denumdoc=canumdoc " & _
         " where dealma='" & Trim(Text2) & "'" & _
         "  and casitgui='V' group by decodigo " & _
         "  order by decodigo"
         
         
  VGCNx.Execute nsql
  
  Bar1.Value = 0
  Label3 = Format(0, "###,##0.00")
  Label4 = Format(0, "###,##0.00")
  DoEvents
  n = 0
  rs.Open "tempora", VGCNx, adOpenDynamic, adLockOptimistic
  If rs.RecordCount > 0 Then
    rs.MoveLast
    Bar1.Max = rs.RecordCount
    Label3 = Format(rs.RecordCount, "###,##0.00")
    VGCNx.BeginTrans
    nflag = 1
    rs.MoveFirst
    Do Until rs.EOF
        nalma = rs!stalma
        ncodi = rs!stcodigo
        
        Set rs1 = VGCNx.Execute("select * from i0 where decodigo='" & ncodi & "'")
        r1 = 0: r2 = 0: r3 = 0
        If rs1.RecordCount > 0 Then
           r1 = rs1!ingreso
           r2 = rs1!Salida
        End If
        rs1.Close
        Set rs1 = Nothing
        r3 = r1 - r2

        VGCNx.Execute "update stkart " & _
                   " set stskdis=" & r3 & _
                   " where stalma='" & nalma & "' and stcodigo='" & ncodi & "'"
        
        n = n + 1
        Label4 = Format(n, "###,##0.00")
        Bar1.Value = n
        DoEvents
        rs.MoveNext
    Loop
    VGCNx.CommitTrans
    nflag = 0
  End If
  rs.Close
  Set rs = Nothing
 
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

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
  Dim rsc As New ADODB.Recordset
  
  Combo1.Clear
  Set rsc = VGCNx.Execute("select TAALMA,TADESCRI from tabalm ")
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
