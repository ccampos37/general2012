VERSION 5.00
Begin VB.Form frmMovAnual 
   Caption         =   "Movimiento Anual por Articulo"
   ClientHeight    =   6030
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   ControlBox      =   0   'False
   LinkTopic       =   "Form12"
   ScaleHeight     =   6030
   ScaleWidth      =   5850
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   4695
      Left            =   240
      TabIndex        =   6
      Top             =   240
      Width           =   5295
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enero"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   63
         Top             =   600
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Febero"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   62
         Top             =   840
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Marzo"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   61
         Top             =   1080
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Abril"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   60
         Top             =   1320
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mayo"
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   59
         Top             =   1560
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Junio"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   58
         Top             =   1800
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Julio"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   57
         Top             =   2040
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Agosto"
         Height          =   255
         Index           =   7
         Left            =   360
         TabIndex        =   56
         Top             =   2280
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Setiembre"
         Height          =   255
         Index           =   8
         Left            =   360
         TabIndex        =   55
         Top             =   2520
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Octubre"
         Height          =   255
         Index           =   9
         Left            =   360
         TabIndex        =   54
         Top             =   2760
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Noviembre"
         Height          =   255
         Index           =   10
         Left            =   360
         TabIndex        =   53
         Top             =   3000
         Width           =   1070
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diciembre"
         Height          =   255
         Index           =   11
         Left            =   360
         TabIndex        =   52
         Top             =   3240
         Width           =   1070
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "600.00"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   51
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "500.00"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   50
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   49
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   48
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   4
         Left            =   1440
         TabIndex        =   47
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   5
         Left            =   1440
         TabIndex        =   46
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   6
         Left            =   1440
         TabIndex        =   45
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   7
         Left            =   1440
         TabIndex        =   44
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   8
         Left            =   1440
         TabIndex        =   43
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   9
         Left            =   1440
         TabIndex        =   42
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   10
         Left            =   1440
         TabIndex        =   41
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   11
         Left            =   1440
         TabIndex        =   40
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "320.00"
         Height          =   255
         Index           =   12
         Left            =   2520
         TabIndex        =   39
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   13
         Left            =   2520
         TabIndex        =   38
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "225.00"
         Height          =   255
         Index           =   14
         Left            =   2520
         TabIndex        =   37
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "15.00"
         Height          =   255
         Index           =   15
         Left            =   2520
         TabIndex        =   36
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   16
         Left            =   2520
         TabIndex        =   35
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   17
         Left            =   2520
         TabIndex        =   34
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   18
         Left            =   2520
         TabIndex        =   33
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   19
         Left            =   2520
         TabIndex        =   32
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   20
         Left            =   2520
         TabIndex        =   31
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   21
         Left            =   2520
         TabIndex        =   30
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   22
         Left            =   2520
         TabIndex        =   29
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   23
         Left            =   2520
         TabIndex        =   28
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "280.00"
         Height          =   255
         Index           =   24
         Left            =   3600
         TabIndex        =   27
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "780.00"
         Height          =   255
         Index           =   25
         Left            =   3600
         TabIndex        =   26
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "555.00"
         Height          =   255
         Index           =   26
         Left            =   3600
         TabIndex        =   25
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "540.00"
         Height          =   255
         Index           =   27
         Left            =   3600
         TabIndex        =   24
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   28
         Left            =   3600
         TabIndex        =   23
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   29
         Left            =   3600
         TabIndex        =   22
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   30
         Left            =   3600
         TabIndex        =   21
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   31
         Left            =   3600
         TabIndex        =   20
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   32
         Left            =   3600
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   33
         Left            =   3600
         TabIndex        =   18
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   34
         Left            =   3600
         TabIndex        =   17
         Top             =   3000
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "--"
         Height          =   255
         Index           =   35
         Left            =   3600
         TabIndex        =   16
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         Caption         =   "    Entrada"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         Caption         =   "     Salida"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   2520
         TabIndex        =   14
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         Caption         =   "   Costo Prom"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   3600
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label10 
         Caption         =   "Promedio Mensual"
         Height          =   495
         Left            =   360
         TabIndex        =   12
         Top             =   3600
         Width           =   855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   11
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   10
         Top             =   3720
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Indice de Rotación   (Salidas/ Año)"
         Height          =   255
         Left            =   975
         TabIndex        =   9
         Top             =   4185
         Width           =   2565
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   3600
         TabIndex        =   8
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         Caption         =   "    Mes"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   735
      Left            =   4560
      Picture         =   "frmMovAnual.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5040
      Width           =   735
   End
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   4920
      Width           =   3855
      Begin VB.Label Label6 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   360
         TabIndex        =   64
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label15 
         Caption         =   "S/."
         Height          =   255
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label14 
         Caption         =   "Ultimo Ingreso"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2040
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmMovAnual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command7_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  central frmMovAnual
  Label12(0) = Format("275", "###0.00")
  Label12(1) = Format("180", "###0.00")
  'Label9(0) = Format("12.20", "###0.00")
  Label9(1) = Format("12.90", "###0.00")
  Label9(2) = Format("12.90", "###0.00")
  Label17 = Format(indice, "###0.00")
  cargarotacion
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set Adoreg1 = Nothing
  Unload Me
End Sub

Private Sub Label12_Click(Index As Integer)
Dim I As Integer
    For I = 0 To 11
       
    Next I
    'Label12(0) = Format("prom0", "###0.00")
   
           
End Sub

Private Sub cargarotacion()
 Dim rsql As String
 Dim Adoreg1 As ADODB.Recordset
 Dim mes As Integer
 Dim sumaent, SumaSal As Double
 Dim nummes As Integer
 For mes = 1 To 12
      Label2(mes - 1).Alignment = 2
      Label2(mes - 1) = "--"
      Label2(12 + mes - 1).Alignment = 2
      Label2(12 + mes - 1) = "--"
      Label2(24 + mes - 1).Alignment = 2
      Label2(24 + mes - 1) = "--"
 Next mes
 Label12(0) = Format(0, "##0.000")
 Label12(1) = Format(0, "##0.000")
 sumaent = 0
 SumaSal = 0
 
 rsql = "select smcodigo,smcanent,smcansal,smmespro,smmnpreuni from moresmes where smalma ='" & VGAlma & "'   and smcodigo = '" & FormMovArt.Text1 & "' and left(smmespro,4)='" & FormMovArt.fecha & "' ORDER BY SMMESPRO "
Set Adoreg1 = New ADODB.Recordset
Adoreg1.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
nummes = Adoreg1.RecordCount
If nummes <> 0 Then
     For mes = 1 To 12
        If Not Adoreg1.EOF Then
            If Adoreg1("smmespro") = Year(Date) & Format(mes, "00") Then
                  Label2(mes - 1).Alignment = 1
                  Label2(mes - 1) = Format(Adoreg1("smcanent"), "##0.00")
                  sumaent = sumaent + Adoreg1("smcanent")
                  Label2(12 + mes - 1).Alignment = 1
                  Label2(12 + mes - 1) = Format(Adoreg1("smcansal"), "###0.00")
                  SumaSal = SumaSal + Adoreg1("smcansal")
                  Label2(24 + mes - 1).Alignment = 1
                  Label2(24 + mes - 1) = IIf(IsNull(Adoreg1("smmnpreuni")), 0, Adoreg1("smmnpreuni"))
                  Label2(24 + mes - 1) = Format(Label2(24 + mes - 1), "#0.00")
                Adoreg1.MoveNext
            End If
        End If
    Next mes
    Label12(0) = sumaent / nummes
    Label12(1) = SumaSal / nummes
    Label12(0) = Format(Label12(0), "##0.000")
    Label12(1) = Format(Label12(1), "##0.000")
End If
 Adoreg1.Close
 
 rsql = "select stkpreult,stkfecult from stkart where stcodigo= '" & FormMovArt.Text1 & "'  and stalma = '" & VGAlma & "' "
 Set Adoreg1 = New ADODB.Recordset
 Adoreg1.Open rsql, VGCNx, adOpenDynamic, adLockOptimistic
 If Adoreg1.RecordCount <> 0 Then
      Label9(2).Caption = Format(Adoreg1("stkpreult"), "###0.000")
      Label9(1) = IIf(IsNull(Adoreg1("stkfecult")), "", Adoreg1("stkfecult"))
 End If
End Sub

Function indice() As Long
Dim adoreg As ADODB.Recordset
Set adoreg = New ADODB.Recordset
Dim fecha As Date

fecha = FormMovArt.fecha
indice = 0
adoreg.Open "Select count(canumdoc) from movalmcab,movalmdet where catd = detd and  detd in ( 'NS','GS') and cafecdoc = '" & fecha & "'" & _
                       " and  canumdoc= denumdoc  and caalma = dealma and dealma ='" & VGAlma & _
                       "' and   decodigo= '" & FormMovArt.Text1 & "'", VGCNx, adOpenStatic
If Not adoreg.EOF Then
   indice = IIf(IsNull(adoreg(0)), 0, adoreg(0))
End If
adoreg.Close
End Function
