VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frCCAR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "C. de Costos y Areas a Considerar"
   ClientHeight    =   5385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frCCAR.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5385
   ScaleWidth      =   3735
   Tag             =   "Panel de Administración de Trabajadores"
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   1972
      TabIndex        =   2
      Top             =   4830
      Width           =   1095
   End
   Begin VB.CommandButton CmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   667
      TabIndex        =   1
      Top             =   4830
      Width           =   1095
   End
   Begin MSComctlLib.TreeView tView 
      Height          =   3480
      Left            =   75
      TabIndex        =   0
      Top             =   1200
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   6138
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   441
      LabelEdit       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      ImageList       =   "Img1"
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Seleccionado por"
      Height          =   960
      Left            =   90
      TabIndex        =   3
      Top             =   105
      Width           =   3585
      Begin VB.OptionButton Option1 
         Caption         =   "&Areas de Trabajo"
         Height          =   210
         Left            =   645
         TabIndex        =   5
         Top             =   345
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton Option2 
         Caption         =   "&Centros de Costo"
         Height          =   210
         Left            =   645
         TabIndex        =   4
         Top             =   615
         Width           =   1605
      End
   End
   Begin MSComctlLib.ImageList Img1 
      Left            =   180
      Top             =   330
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCCAR.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCCAR.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCCAR.frx":14C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCCAR.frx":1912
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCCAR.frx":1C2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frCCAR.frx":2506
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frCCAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Aux As Boolean
Dim KEYAUX As String
Private Sub CMDACEPTAR_CLICK()
    Dim XNODE As NODE
    VPTRASPRM = ""
    For Each XNODE In tView.Nodes
        If XNODE.Checked Then VPTRASPRM = VPTRASPRM & ",'" & Right(XNODE.KEY, Len(XNODE.KEY) - 1) & "'"
    Next
    If VPTRASPRM <> "" Then VPTRASPRM = "('X'" & VPTRASPRM & ")"
    If Option1.Value Then VPNUMTMP = 1 Else VPNUMTMP = 2
    Unload Me
End Sub

Private Sub Command2_Click()
    VPTRASPRM = ""
    Unload Me
End Sub

Private Sub Form_Activate()
    Option1.Value = True
    Option1.Value = False
    Option2.Value = True
    Option1.Value = True
End Sub

Private Sub Form_Load()
    CARGAARBOL
    Dim NODO As NODE
    For Each NODO In tView.Nodes
        NODO.Checked = True
    Next NODO
End Sub

Public Sub CARGAARBOL()
    Dim ICONRAIZ As Integer, ICONSUB As Integer, ICONINDICE As Integer
    Dim ICONAUX As Integer
    tView.Nodes.Clear
    Dim RSCCOSTO As New ADODB.Recordset
    INTO = ""
    If Option1.Value Then 'SI ES POR AREAS
        RSCCOSTO.Open "SELECT *  FROM AREASTRAB ORDER BY CODCCOSTO", DBSYSTEM, adOpenKeyset, adLockOptimistic
        ICONRAIZ = 3
        ICONINDICE = 6
        ICONSUB = 1
        
    Else
        RSCCOSTO.Open "SELECT *  FROM CCOSTOS ORDER BY CODCCOSTO", DBSYSTEM, adOpenKeyset, adLockOptimistic
        ICONRAIZ = 4
        ICONINDICE = 5
        ICONSUB = 2
        
    End If
    If RSCCOSTO.RecordCount = 0 Then
        Set RSCCOSTO = Nothing
        MsgBox "EL ARCHIVO DE AREAS DE TRABAJO O CENTROS DE COSTOS SE ENCUENTRA VACÍA, POR FAVOR INICIALIZE LOS CENTROS DE COSTO", vbCritical
        Exit Sub
    End If
    Dim XNODE As NODE
    Dim COD As String
    Set XNODE = tView.Nodes.Add(, , "RAIZ", IIf(Option1.Value, "TODAS LAS AREAS DE TRABAJO", "TODOS LOS CENTROS DE COSTO"), ICONRAIZ)
        XNODE.Checked = True
    With RSCCOSTO
        Do While Not .EOF
            If Len(!CODCCOSTO) = 2 Or InStr(!CODCCOSTO, ".") = 0 Then
                COD = "RAIZ"
                ICONAUX = ICONINDICE
            Else
               X = 1
               Do While Not X = 0
                  X = InStr(X + 1, !CODCCOSTO, ".")
                  If X <> 0 Then Y = X
               Loop
               COD = "C" & Mid(!CODCCOSTO, 1, Y - 1)
               ICONAUX = ICONSUB
            End If
             Set XNODE = tView.Nodes.Add(COD, 4, "C" & !CODCCOSTO, !NOMBRE, ICONAUX)
                 XNODE.Checked = True
                 XNODE.Tag = COD
            .MoveNext
        Loop
    End With
    tView.Nodes("RAIZ").Expanded = True
    tView.Nodes("RAIZ").Selected = True
    Set RSCCOSTO = Nothing
End Sub

Private Sub OPTION1_Click()
    CARGAARBOL
End Sub

Private Sub OPTION2_Click()
    CARGAARBOL
End Sub
Private Sub TVIEW_NODECHECK(ByVal NODE As MSComctlLib.NODE)
Dim NODO As NODE
Dim FLAG As Boolean
Dim FLAG2 As Boolean
Dim CHK As Boolean
Aux = False
If NODE.Checked Then
    CHK = True
  Else: CHK = False
End If
FLAG2 = False
KEY = NODE.KEY
If NODE.Children = 0 Then Exit Sub
KEYAUX = ""
RECURS NODE, NODE.Child.LastSibling.KEY
If KEYAUX = "" Then KEYAUX = NODE.Child.LastSibling.KEY
    For Each NODO In tView.Nodes
        If NODO.Tag = KEY And FLAG2 = False Then FLAG = True
            If FLAG = True Then
                FLAG2 = True
                NODO.Checked = CHK
                If NODE.Children > 0 Then
                    If NODO.KEY = KEYAUX Then Exit For
                End If
            End If
    Next
End Sub
Private Function RECURS(NODO As NODE, KEY As String) As String
Dim FLAG As Boolean
Dim FLAG2 As Boolean
Dim NODOAUX As NODE
FLAG2 = False
FLAG = False
    For Each NODOAUX In tView.Nodes
        If Aux = True Then Exit For
        If NODOAUX.Tag = KEY And FLAG2 = False Then FLAG = True
        If FLAG = True Then
            FLAG2 = True
            If NODOAUX.LastSibling.Children > 0 Then
                RECURS NODOAUX.LastSibling, NODOAUX.LastSibling.KEY
              Else
                KEYAUX = NODOAUX.LastSibling.KEY
                Aux = True
                Exit For
            End If
        End If
    Next
End Function


