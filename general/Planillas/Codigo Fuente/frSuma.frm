VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frSuma 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resultados de Suma"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5925
   Icon            =   "frSuma.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView Vista 
      Height          =   2745
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   4842
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Codigo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Total"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   660
      Top             =   1980
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
            Picture         =   "frSuma.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frSuma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim RSCNPT As New ADODB.Recordset
    Dim RSSUM As New ADODB.Recordset
    Dim XLIST As ListItem, X As Integer
    Select Case UCase(VPTAREA)
        Case "INPUT"
            RSCNPT.Open "CONCEPTOS", DBSYSTEM, adOpenKeyset, adLockOptimistic
            For X = 2 To InputPl.dgInput.Columns.Count - 1
                RSCNPT.MoveFirst
                RSCNPT.FIND "CODIGO='" & InputPl.dgInput.Columns(X).DataField & "'"
                Set XLIST = Vista.ListItems.Add(, RSCNPT!Codigo, RSCNPT!Codigo, , 1)
                XLIST.SubItems(1) = RSCNPT!NOMBRE
                RSSUM.Open "SELECT SUM(" & XLIST.KEY & ") AS TOTAL FROM INPUTBOL WHERE " & XLIST.KEY & "<>NULL", DBSYSTEM, adOpenStatic
                XLIST.SubItems(2) = Format(RSSUM!TOTAL, "0.00")
                RSSUM.Close
            Next
            Vista.Visible = True
        Case "PLANILLAS"
            RSCNPT.Open "SELECT * FROM COLUMPL ORDER BY INDICE", DBSYSTEM, adOpenStatic
            Do While Not RSCNPT.EOF
                Set XLIST = Vista.ListItems.Add(, RSCNPT!Codigo, RSCNPT!Codigo, , 1)
                XLIST.SubItems(1) = RSCNPT!NOMBRE
                RSSUM.Open "SELECT SUM(" & XLIST.KEY & ") AS TOTAL FROM " & REGSISTEMA.TABLAPLAN & " WHERE MES=" & DateSQL(CDate(frPlans.LPlans.SelectedItem.SubItems(1))), DBSYSTEM, adOpenKeyset, adLockOptimistic
                XLIST.SubItems(2) = Format(RSSUM!TOTAL, "##,##0.00")
                RSSUM.Close
                RSCNPT.MoveNext
            Loop
            Vista.Visible = True
        Case "FRINPUTMOV"
            RSCNPT.Open "SELECT * FROM CONCEPTOS WHERE TIPO<>0 AND ESESCRITO=1 ORDER BY FILA", DBSYSTEM, adOpenStatic
            If RSCNPT.RecordCount = 0 Then
                MsgBox "NO EXISTEN DATOS POR MOSTRAR", vbInformation
                Set RSCNPT = Nothing
                Exit Sub
            End If
            For X = 2 To frInputMov.DGLista.Columns.Count - 1
                RSCNPT.MoveFirst
                RSCNPT.FIND "CODIGO='" & Trim$(frInputMov.DGLista.Columns(X).DataField) & "'"
                If Not RSCNPT.EOF Then
                    Set XLIST = Vista.ListItems.Add(, RSCNPT!Codigo, RSCNPT!Codigo, , 1)
                    XLIST.SubItems(1) = RSCNPT!NOMBRE
                    RSSUM.Open "SELECT SUM(" & XLIST.KEY & ") AS TOTAL FROM  [##INPUTMOV" & VGL_COMPUTER & "] ", DBSTARPLAN, adOpenKeyset, adLockOptimistic
                    XLIST.SubItems(2) = Format(RSSUM!TOTAL, "##,##0.00")
                    RSSUM.Close
                End If
            Next
            Vista.Visible = True
        Case Else
            MsgBox "NO SE HA CARGADO LAS SUMAS PARA EL PROCEDIMIENTO INDICADO"
    End Select
    Set RSCNPT = Nothing
    Set RSSUM = Nothing
End Sub

