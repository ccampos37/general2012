VERSION 5.00
Begin VB.Form FormProValCierre 
   Caption         =   "Cierre Mensual de Valorizacion"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   LinkTopic       =   "Form1"
   ScaleHeight     =   4290
   ScaleWidth      =   5100
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      Height          =   1095
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   4215
      Begin VB.OptionButton Option2 
         Caption         =   "Definitivo"
         Height          =   195
         Left            =   480
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Previo"
         Height          =   255
         Left            =   480
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   735
      Left            =   1320
      Picture         =   "FormAyu.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3360
      Width           =   735
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   2640
      Picture         =   "FormAyu.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3360
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Valorizacion  de Almacenes"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   1440
      Width           =   4215
      Begin VB.Label Label3 
         Caption         =   "Mes"
         Height          =   255
         Left            =   2520
         TabIndex        =   8
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Valorizar al  dia"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label1 
         Caption         =   "Cierre Anterior"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FormProValCierre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mespro As String
''Dim db As Database

Private Sub Command1_Click()
Dim rsql As String
Dim rs As Recordset
  mespro = "199912"
  valorizacion
   'usql = "insert into MoResMes (SMCIA,SMALMA,SMCODIGO,SMMESPRO,SMMNPRE) VALUES (,'" & VGAlma & "','" & cadena & "','" & mespro & "' ,'" & CANT & "','" & cantdolar & "') "
'  Rsql = "select SMCODIGO  from MoResMes  where  SMALMA = '" & VGAlma & "' AND SMMESPRO='" & mespro & "'"
'  Set rs = Db.OpenRecordset(Rsql, dbOpenSnapshot)
'  While Not rs.EOF
'      finmes (rs(0))
'      rs.MoveNext
'  Wend
End Sub

Private Sub Command8_Click()
 Unload Me
 
End Sub

Private Sub valorizacion()
Dim Aux As String
'Dim aux As Integer
 
 Dim rs As Recordset
 Dim Rsql1 As String
 Dim dato1 As Date
 Dim dato2 As Date
 Dim tipo As String
 Dim cero As Boolean
 ' dato1 = "12/01/1999" 'Format("01/12/1999", "mm/dd/yyyy")
 'dato2 = "12/30/1999" 'Format("30/12/1999", "mm/dd/yyyy")
 tipo = "NI"
 Rsql1 = "select n.CANUMDOC  from MovAlmCab n where  n.CAALMA ='" & VGAlma & "' AND n.CATD = 'NI' AND n.CAFECDOC >= #" & dato1 & "#   AND n.CAFECDOC < #" & dato2 & "#" '
 'Set db = Workspaces(0).OpenDatabase(cRuta2)
 'Set RS = db.OpenRecordset(Rsql1, dbOpenDynaset)
 
 Set rs = VGCNx.Execute(Rsql1)
 While Not rs.EOF
     Call buscarDoc(tipo, rs(0), cero)
     If cero Then
         MsgBox "No se puede valorizar,Tienes que valorizar articulos pendientes", vbCritical, "Aviso"
         rs.Close
         Exit Sub
     End If
     rs.MoveNext
 Wend
 
 Rsql1 = "select SMCODIGO  from MoResMes  where  SMALMA = '" & VGAlma & "' AND SMMESPRO='" & mespro & "'"
' Set RS = db.OpenRecordset(Rsql1, dbOpenSnapshot)
 Set rs = VGCNx.Execute(Rsql1)
 While Not rs.EOF
      finmes (rs(0))
      rs.MoveNext
 Wend
 rs.Close
 MsgBox "El proceso de  valorizacion Finalizado", vbInformation, "Aviso"
'db.Close

End Sub

Public Sub buscarDoc(doc As String, NumDoc As String, cero As Boolean)
  Dim rs As Recordset
  Dim rsql As String   'PRECIO
  cero = False
  rsql = "select  DECODIGO, DEPRECIO from MovAlmDet   where DEALMA = '" & VGAlma & "'  and DETD= '" & doc & "'  AND  DENUMDOC= '" & NumDoc & "'"
  'Set db = Workspaces(0).OpenDatabase(cRuta2)
  'Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
  
  Set rs = VGCNx.Execute(rsql)
  While Not rs.EOF
            If rs(1) = 0 Then
              cero = True
              rs.Close
              Exit Sub
            End If
            rs.MoveNext
 Wend
End Sub

Private Sub finmes(CADENA As String)
  Dim cant As Double
  Dim cantdolar As Double
  Dim rsql As String
  Dim uSql As String
  Dim rs As Recordset
  Dim rs1 As Recordset
  
     mespro = Year(Date) & Format(Month(Date), "00")
     rsql = "select STKPREPRO from stkart where  STALMA = '" & VGAlma & "' AND STCODIGO='" & CADENA & "'"  '
    ' Set db = Workspaces(0).OpenDatabase(cRuta2)
    ' Set RS = db.OpenRecordset(RSQL, dbOpenSnapshot)
    
    Set rs = VGCNx.Execute(rsql)
    
     If IsNull(rs(0)) Then
         MsgBox "no hay precio de compra", vbCritical, "Error"
         cant = 0
       Else
         cant = rs(0)
       End If
       cantdolar = cant * VGTipCamb '* tipo(Fecha)
       'Update de moresmes ya que cada entrada ha sido
       uSql = "update MoResMes set SMMNPREUNI = " & cant & "   where  SMALMA = '" & VGAlma & "' AND SMMESPRO='" & mespro & "' and SMCODIGO='" & CADENA & "' "
'       usql = "update MoResMes set SMUSPREUNI = " & cantdolar & "   where SMCIA = '" & VGCODEMPRESA & "' AND SMALMA = '" & VGAlma & "' AND SMMESPRO='" & mespro & "' and SMCODIGO='" & cadena & "' "
'       usql = "update MoResMes set SMMNENT = " & CANT & "   where SMCIA = '" & VGCODEMPRESA & "' AND SMALMA = '" & VGAlma & "' AND SMMESPRO='" & mespro & "' and SMCODIGO='" & cadena & "' "
'       usql = "update MoResMes set SMMNSAL = " & CANT & "   where SMCIA = '" & VGCODEMPRESA & "' AND SMALMA = '" & VGAlma & "' AND SMMESPRO='" & mespro & "' and SMCODIGO='" & cadena & "' "
'       usql = "update MoResMes set SMUSENT = " & CANT & "   where SMCIA = '" & VGCODEMPRESA & "' AND SMALMA = '" & VGAlma & "' AND SMMESPRO='" & mespro & "' and SMCODIGO='" & cadena & "' "
'       usql = "update MoResMes set SMMNSAL = " & CANT & "   where SMCIA = '" & VGCODEMPRESA & "' AND SMALMA = '" & VGAlma & "' AND SMMESPRO='" & mespro & "' and SMCODIGO='" & cadena & "' "
       
       VGCNx.Execute uSql
       'rs.MoveNext
   ' Wend
End Sub

Private Sub cerrar()
'    COPY TO &WFILE ALL
'      Close Data
'      USE MORESME
'      COPY TO SALVAME
'      Close Data
'      USE SALVAME
'      INDEX ON SMALMA+SMCODIGO TO SALVAME1
'      Close Data
'
'      IF FILE ('SALVANT.DBF')
'         USE SALVANT
'         IF EOF ()
'            WFILE2= 'SKVA0000'
'         Else
'            WFILE2= 'SKVA'+SUBSTR(WFILE,5,4)
'         End If
'         COPY TO &WFILE2 ALL
'         Close Data
'         USE SALVACT
'         COPY TO SALVANT
'         Close Data
'         USE SALVANT
'         INDEX ON SMVALMA + SMVCODIGO TO SALVANT1
'         Close Data
'      End If
'
'      AMENSA={'Proceso de Valorizacion Finalizado'}
'      FG_AMENSAJE(2, AMENSA,,PU_CINVER)
'  Else
'      MENSAJE('Hay Articulos sin Valorizar no se puede Cerrar el Mes',1)
'  End If
'End If
End Sub

