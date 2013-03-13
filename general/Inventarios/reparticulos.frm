VERSION 5.00
Begin VB.Form FormArtRep 
   Caption         =   "Reporte de Articulo"
   ClientHeight    =   3720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4530
   LinkTopic       =   "Form12"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   4530
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3855
      Begin VB.OptionButton Option6 
         Caption         =   "Por Descripcion"
         Height          =   375
         Left            =   600
         TabIndex        =   7
         Top             =   1080
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Por Casillero"
         Height          =   255
         Left            =   1800
         TabIndex        =   6
         Top             =   1560
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Por Familia"
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   1560
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Por Codigo Fabricante"
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Codigo"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2520
      Picture         =   "reparticulos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1080
      Picture         =   "reparticulos.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   855
   End
   Begin VB.Frame FrameRep 
      Height          =   2295
      Left            =   480
      TabIndex        =   8
      Top             =   240
      Width           =   3495
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   12
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   11
         Top             =   1080
         Width           =   1455
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Todos los articulos"
         Height          =   375
         Left            =   360
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Rango"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Fin"
         Height          =   255
         Left            =   360
         TabIndex        =   14
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "InicIo "
         Height          =   255
         Left            =   360
         TabIndex        =   13
         Top             =   1080
         Width           =   735
      End
   End
End
Attribute VB_Name = "FormArtRep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim flag As Boolean
Dim db As Database
Private Sub Command1_Click()
   If Option1.Value Then
          artcod
   ElseIf Option2.Value Then
          artcod1
   ElseIf Option3.Value Then
          artfam
'   ElseIf Option4.Value Then
'          artgrupo
   ElseIf Option5.Value Then
          artcas
   Else
          artdes
   End If
   limpiar_t1_t2
   Frame1.Visible = True
End Sub


Private Sub artcod()

  Dim cadena As String
           If Option7.Value Then
               'If Text1 <> "" Then Exit Sub
               formrep.CrystalReport1.ReportFileName = RUTA & "reporteInv\catartcod.rpt"
                Ubi_Tab formrep.CrystalReport1
               formrep.CrystalReport1.DiscardSavedData = True
               If Text2 <> "" Then    '  "23134671"
                        cadena = "{MAEART.ACODIGO} in '" & Text1 & "' to '" & Text2 & "'"
                Else
                        cadena = "{MAEART.ACODIGO} = '" & Text1 & "' "
                End If
                formrep.CrystalReport1.SelectionFormula = cadena
            Else
               formrep.CrystalReport1.ReportFileName = RUTA & "reporteInv\catartcod.rpt"
               Ubi_Tab formrep.CrystalReport1
               formrep.CrystalReport1.Formulas(0) = "empresa ='" & VGNemp & "'"
               formrep.CrystalReport1.DiscardSavedData = True
            End If
            formrep.CrystalReport1.Destination = crptToWindow
            'formrep.CrystalReport1.WindowTitle = "Impresión de Catalogo de Articulos"
            If formrep.CrystalReport1.Status <> 2 Then
               formrep.CrystalReport1.Action = 1
            End If
End Sub

Private Sub artcod1()

            
               formrep.CrystalReport1.ReportFileName = RUTA & "reporteInv\catartcod2.rpt"
               Ubi_Tab formrep.CrystalReport1
               formrep.CrystalReport1.Formulas(0) = "empresa ='" & VGNemp & "'"
               formrep.CrystalReport1.DiscardSavedData = True
               formrep.CrystalReport1.Destination = crptToWindow
            'frmrep.CrystalReport1.SelectionFormula = Formula_Impresion$
            'formrep.CrystalReport1.WindowTitle = "Impresión de Catalogo de Articulos"
            formrep.CrystalReport1.Action = 1
End Sub
Private Sub artdes()
       formrep.CrystalReport1.ReportFileName = RUTA & "reporteInv\catartdescp.rpt"
       Ubi_Tab formrep.CrystalReport1
       formrep.CrystalReport1.DiscardSavedData = True
       formrep.CrystalReport1.Formulas(0) = "empresa ='" & VGNemp & "'"
       formrep.CrystalReport1.Destination = crptToWindow
            'frmrep.CrystalReport1.SelectionFormula = Formula_Impresion$
       formrep.CrystalReport1.Action = 1
End Sub

Private Sub artfam()
       formrep.CrystalReport1.ReportFileName = RUTA & "reporteInv\catfam.rpt"
       Ubi_Tab formrep.CrystalReport1
       formrep.CrystalReport1.DiscardSavedData = True
        formrep.CrystalReport1.Formulas(0) = "empresa ='" & VGNemp & "'"
       formrep.CrystalReport1.Destination = crptToWindow
            'frmrep.CrystalReport1.SelectionFormula = Formula_Impresion$
     
       formrep.CrystalReport1.Action = 1
End Sub

Private Sub artgrupo()

End Sub
   
Private Sub artcas()
       formrep.CrystalReport1.ReportFileName = RUTA & "reporteInv\artcas.rpt"
        Ubi_Tab formrep.CrystalReport1
       formrep.CrystalReport1.DiscardSavedData = True
       formrep.CrystalReport1.Destination = crptToWindow
            'frmrep.CrystalReport1.SelectionFormula = Formula_Impresion$
       formrep.CrystalReport1.Action = 1

End Sub
   
Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub Form_Load()
   central Me
   flag = False
   Option1.Value = True
   flag = True
End Sub

Private Sub Option1_Click()
 If flag Then
  Frame1.Visible = False
  FrameRep.Visible = True
 End If
End Sub

Private Sub Text1_DblClick()
   VGForm1 = 7
   FormAyuArt1.Show 1
   If Text1 <> "" And Text2 <> "" And Text1 > Text2 Then
      MsgBox "Ingrese un codigo menor al fin ", vbOKOnly, "Error"
      Exit Sub
   End If
   If Text1 <> "" Then
        Text2.Enabled = True
        Text2.SetFocus
   End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 And Text1 <> "" Then
      If Existe_cod_art(Text1) <> "" Then
              Text2.Enabled = True
              Text2.SetFocus
      End If
   End If
End Sub

Private Sub Text2_DblClick()
   FormAyuArt1.Show 1
   If Text2 <> "" Then
        Command1.SetFocus
   End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 And Text2 <> "" Then
      
      If Existe_cod_art(Text2) <> "" Then
         If Text1 > Text2 Then
         
         MsgBox "El codigo fin debe ser mayor que el inicio"
                               Exit Sub
         End If
         Command1.SetFocus
      End If
      
   End If
End Sub

Function Existe_cod_art(text As TextBox) As String
 
 Dim rs As Recordset
 Dim rsql As String
  rsql = "select  ACODIGO FROM maeart where ACODIGO = '" & text & "'" '
  Set db = Workspaces(0).OpenDatabase(RUTA & NAMEBD)
  Set rs = db.OpenRecordset(rsql, dbOpenSnapshot)
  If Not rs.EOF Then
    Existe_cod_art = rs(0)
  Else
    MsgBox "El tipo de codigo no existe !", vbOKOnly, "Error"
    Existe_cod_art = ""
  End If
   rs.Close
End Function

Private Sub limpiar_t1_t2()
  Text1 = ""
  Text2 = ""
 
End Sub
