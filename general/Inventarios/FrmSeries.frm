VERSION 5.00
Begin VB.Form FrmSeries 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Series"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4110
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4110
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   405
      Left            =   1140
      TabIndex        =   2
      Top             =   2580
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Height          =   2355
      Left            =   255
      TabIndex        =   0
      Top             =   75
      Width           =   3510
      Begin VB.ListBox List1 
         Height          =   1620
         Left            =   360
         TabIndex        =   1
         Top             =   285
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FrmSeries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim adodc1 As ADODB.Recordset
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim criterio As String
Set adodc1 = New ADODB.Recordset
criterio = "select * from stkseri  where  STSALMA = '" & VGAlma & "' AND STSCODIGO = '" & VGcod & " ' and STSSKDIS <> 0 "
adodc1.Open criterio, VGcnx, adOpenStatic
If adodc1.RecordCount > 0 Then
  Frame1.Caption = UCase(VGcod)
  While Not adodc1.EOF
    List1.AddItem adodc1("stsserie")
    adodc1.MoveNext
  Wend

End If

adodc1.Close
End Sub
