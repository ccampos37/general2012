VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmTransMABEL 
   Caption         =   "Enviar de IASA a Mabel"
   ClientHeight    =   5025
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9585
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   5025
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "&Enviar"
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Agrupar"
      Height          =   4335
      Left            =   5040
      TabIndex        =   14
      Top             =   120
      Width           =   4455
      Begin VB.CommandButton CmdBuscarArt 
         Caption         =   "..."
         Height          =   255
         Left            =   3360
         TabIndex        =   7
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1320
         TabIndex        =   6
         Top             =   240
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid Salida 
         Height          =   3015
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   5318
         _Version        =   393216
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label4 
         Caption         =   "Descripcion"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Codigo"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8280
      TabIndex        =   11
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6960
      TabIndex        =   10
      Top             =   4560
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Guia"
      Height          =   4335
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4935
      Begin VB.TextBox TxtNroGuia 
         Height          =   285
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   240
         Width           =   3135
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3015
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   5318
         _Version        =   393216
         SelectionMode   =   1
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   600
         Width           =   1335
         _ExtentX        =   1879
         _ExtentY        =   370
         _Version        =   393216
         Format          =   24510465
         CurrentDate     =   36686
      End
      Begin VB.Label Label8 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label7 
         Caption         =   "Nro. Guia"
         Height          =   255
         Left            =   2640
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "<< Selecione con la tecla TAB o un click del Mouse >>"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   960
         Width           =   4455
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicacion"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmTransMABEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim promedio As Double
Dim CANTIDAD As Double
Dim Adoreg1 As ADODB.Recordset
Dim adoreg4 As ADODB.Recordset
Dim nroguia As String
Dim cantidadxCodigo  As Double
Dim nSalida As Double
Dim cNroTxGuia As String

Private Sub CmdBuscarArt_Click()
  VGForm1 = 12
  FormAyuArt1.Show 1
End Sub

Private Sub CmdGrabar_Click()
Dim i As Integer, f As Integer
 If Trim(Text1) = "" Then
     MsgBox "Antes de Grabar ,debe asociar", vbExclamation, "Guia"
     Exit Sub
 End If
 If Salida.Rows = 1 Then Exit Sub
 Call grabar
 nSalida = 0
otro:
 f = FG1.Rows - 1
  For i = 0 To f
     If FG1.TextMatrix(i, 0) = ">>" Then
          If FG1.Rows > 2 Then
              FG1.RemoveItem i
              GoTo otro
          Else
              FG1.Clear
              FG1.Rows = 1
              FG1.Row = 0
              inicializaFG1
              Exit For
          End If
     End If
  Next i
 
  Salida.Clear
  Salida.Rows = 1
  InicializaSalida
 ' limpiaTexto
  Text1 = ""
  Label3 = ""
End Sub

Private Sub Command3_Click()
 On Error GoTo cErrorAbrir
 CommonDialog1.CancelError = True
 CommonDialog1.Filter = "Dos (*.dbf)|*.dbf|Todos los Archivos (*.*)|*.*"
 CommonDialog1.FilterIndex = 1
 CommonDialog1.Action = 1
 Text11 = CommonDialog1.FileName
SalirAbrir:
 Text11.SetFocus
 Exit Sub
cErrorAbrir:
 Resume SalirAbrir
End Sub
'Salir del Formulario
Private Sub Command4_Click()
If nSalida = 1 Then
    cConexCom.Execute "Delete From GuiaMabel Where nroguia = '" & nroguia & "' and giasa= '" & cNroTxGuia & "'"
Else
    'nroguia = val(nroguia) + 1
    If Trim(nroguia) <> "" Then
        cConexCom.Execute "update tabalm  set tanument ='" & nroguia & "' where taalma= '" & VGAlma & "' "
    End If
End If
Unload Me
End Sub
'Enviar
Private Sub command5_Click()
   Dim i As Integer
   If FG1.Rows = 1 Then Exit Sub
   Salida.Rows = 1
   For i = 0 To FG1.Rows - 1
      If FG1.TextMatrix(i, 0) = ">>" Then
        Salida.AddItem (FG1.TextMatrix(i, 1) & vbTab & FG1.TextMatrix(i, 2) & vbTab & FG1.TextMatrix(i, 3) & vbTab & FG1.TextMatrix(i, 4) & vbTab & FG1.TextMatrix(i, 5) & vbTab & FG1.TextMatrix(i, 6))
      End If
   Next i
End Sub

Private Sub DTPicker1_Change()
        VGTipCamb = DevolverTCambio(DTPicker1.Value)
End Sub

Private Sub Form_Load()
'  Data3.DatabaseName = cRuta2
'  Data3.RecordSource = "Select * from stkart"
  central Me
  FG1.Clear
  Salida.Clear
  inicializaFG1
  InicializaSalida
  limpiaTexto
  DTPicker1 = Date
  VGTipCamb = DevolverTCambio(DTPicker1.Value)

  nSalida = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        VGTipCamb = DevolverTCambio(VG_FecTrab)
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Existe(1, Text1, "MaeArt", "acodigo", False) = False Then
        MsgBox "El Código del Artículo no existe", vbInformation, "Información"
        Text1.SetFocus
    Else
        Label3 = Devolver_Dato(1, Text1, "MaeArt", "acodigo", False, "adescri")
    End If
End If
End Sub

Private Sub Text11_DblClick()
 If InStr(Trim(Text11), "dbf") Then
      TxtNroGuia.SetFocus
 Else
   MsgBox "Se debe seleccionar un archivo con extension g**.dbf", vbExclamation, "Error"
 End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
  'Dim db As Database
  Dim rs As Recordset
  Dim rutalectura As String
  Dim nombretabla As String
  Dim RSQL As String
  Dim numero As Long
  On Error GoTo Err
   nroguia = ""
   If KeyAscii = 13 And InStr(Trim(Text11), ".dbf") Then
           If Trim(TxtNroGuia) = "" Then
                MsgBox "Ingrese numero de Guia", vbInformation, "Informacion"
                TxtNroGuia.SetFocus: Exit Sub
           Else
                RSQL = "select nroguia from guiamabel where giasa= '" & TxtNroGuia & "' "
                Set adoreg4 = New ADODB.Recordset
                adoreg4.Open RSQL, cConexCom, adOpenStatic
                If Not adoreg4.EOF Then
                    MsgBox "Ya se registro el documento", vbExclamation, "Aviso"
                    adoreg4.Close:     Exit Sub
                End If
           End If
           If InStr(UCase(Trim(Text11)), "G") = 0 Then
              MsgBox "Debe seleccionar una guia de remisión", vbExclamation, "Aviso"
              Enfoque Text11
              Text11.SetFocus
              Exit Sub
           End If
           Text11 = Trim(Text11)
           nombretabla = Mid(Text11, InStrRev(Text11, "\") + 1)
           nombretabla = Left(nombretabla, InStrRev(nombretabla, ".") - 1)
           rutalectura = Left(Text11, InStr(Text11, ":"))
           RSQL = "select  TANUMENT from TabAlm  WHERE TAALMA='" & VGAlma & "' "
           'Set RS = VGBaseDatos.OpenRecordset(RSQL, dbOpenSnapshot)
           Set rs = cConexCom.Execute(RSQL)
           If rs.RecordCount > 0 Then
                nroguia = Format(rs(0) + 1, "0000000000")
           End If
           rs.Close
'           Set db = Workspaces(0).OpenDatabase(rutalectura, False, False, "FoxPro 2.6")        '     "C:\WINDOWS
'           Set RS = db.OpenRecordset(nombretabla, dbOpenDynaset)
           Set rs = cConexCom.Execute("select * from " & nombretabla)
           If rs.EOF Then
              MsgBox "No hay registro  en la guia", vbCritical
              End
           End If
           FG1.Rows = 1
           rs.MoveFirst
           If rs("guia") <> TxtNroGuia Then
              MsgBox "Ingrese el numero de guia correcto", vbExclamation, "Aviso"
              rs.Close
              Exit Sub
           End If
           
           While Not rs.EOF
               If existe_codigo_mabel(rs(0)) Then   'si existe el codigo no  entra al flex
                    FG1.AddItem (" " & vbTab & rs(0) & vbTab & rs(1) & vbTab & rs(2) & vbTab & Format(rs(3), "####0.00") & vbTab & Format(rs(4), "####0.000") & vbTab & rs(5))
               Else
                    cNroTxGuia = TxtNroGuia
                    Set Adoreg1 = New ADODB.Recordset
                    RSQL = "select * from guiamabel"
                    Adoreg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
                    Adoreg1.AddNew
                    Adoreg1("nroguia") = nroguia
                    Adoreg1("acodigo") = Text1.text
                    Adoreg1("codigo") = rs(0)
                    Adoreg1("descri") = rs(1)
                    Adoreg1("unidad") = rs(2)
                    Adoreg1("cantidad") = rs(3)
                    Adoreg1("precio") = rs(4)
                    Adoreg1("giasa") = rs(5)
                    Adoreg1.UpdateBatch
                    nSalida = 1
               End If
               rs.MoveNext
           Wend
           rs.Close
   End If
   Exit Sub
Err:
    numero = Err.Number
    If -2147467259 = Err.Number Then
        MsgBox "Ya se registro el documento", vbExclamation, "Aviso"
   Else
       MsgBox Err.Description
    End If
End Sub

Private Sub FG1_Click()
  If FG1.Row = 0 Then
    Exit Sub
  Else
    If FG1.TextMatrix(FG1.Row, 0) = ">>" Then
         FG1.TextMatrix(FG1.Row, 0) = " "
    Else
         FG1.TextMatrix(FG1.Row, 0) = ">>"
    End If
 End If
End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)
   If KeyAscii = vbKeySpace Then
        FG1_Click
   End If
End Sub

Private Sub grabar()
   Dim i As Integer
   Dim CANT As Double
   Dim monto As Double
   Dim RSQL As String
   Set Adoreg1 = New ADODB.Recordset
   'Dim rsql As String
   On Error GoTo Err
  
   RSQL = "select * from guiamabel"
   Adoreg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
   For i = 1 To Salida.Rows - 1
         If existe_codigo_mabel(Salida.TextMatrix(i, 0)) Then   'si existe el codigo no  entra al flex
                Adoreg1.AddNew
                Adoreg1("nroguia") = nroguia
                Adoreg1("acodigo") = Text1
                Adoreg1("codigo") = Salida.TextMatrix(i, 0)
                Adoreg1("descri") = Salida.TextMatrix(i, 1)
                Adoreg1("unidad") = Salida.TextMatrix(i, 2)
                Adoreg1("cantidad") = Salida.TextMatrix(i, 3)
                Adoreg1("precio") = Salida.TextMatrix(i, 4)
                Adoreg1("giasa") = Salida.TextMatrix(i, 5)
                Adoreg1.UpdateBatch
                Adoreg1.Requery
         Else
                CANT = Salida.TextMatrix(i, 3) + cantidadxCodigo
                RSQL = "Update guiamabel set cantidad =" & CANT & " where   nroguia='" & nroguia & "' and  acodigo='" & Text1 & "' and  codigo ='" & Salida.TextMatrix(i, 0) & "'and  giasa ='" & Salida.TextMatrix(i, 5) & "'"
                cConexCom.Execute RSQL
         End If
         monto = monto + Salida.TextMatrix(i, 4) * Salida.TextMatrix(i, 3)
   Next i
  
   cConexCom.Execute "update tabalm  set tanument ='" & nroguia & "' where taalma= '" & VGAlma & "' "
   
   If frmTransMABEL.Caption = "Recepción de Guia IASA a Mabel" Then
       InicializaSalida
       'limpiaTexto
       Exit Sub
   End If
   If monto <> 0 Then
       promedio = CANT / monto
       CANTIDAD = CANT
       'grabastk
       'grabacabecera
       'grabadetalle
       InicializaSalida
       'limpiaTexto
   Else
       MsgBox "El monto no puede ser igual a cero ", vbExclamation, "Grabar"
   End If
   Set Adoreg1 = Nothing
   Exit Sub
Err:
 monto = Err.Number
  MsgBox Err.Description
  'MsgBox err.Number
End Sub

Private Sub inicializaFG1()
  FG1.FormatString = "^Seleccion | Codigo Art.|Descripcion |^Unidad|Cantidad | Precio Unit| "
  FG1.Row = 0
  FG1.Cols = 7
  FG1.ColWidth(0) = 900
  FG1.ColWidth(1) = 1200
  FG1.ColWidth(2) = 2500
  FG1.ColWidth(3) = 890
  FG1.ColWidth(4) = 1200
  FG1.ColWidth(5) = 1200
  FG1.ColWidth(6) = 2
End Sub

Private Sub InicializaSalida()
  Salida.FormatString = "    Codigo Art.|   Descripcion |  Unidad|Cantidad | Precio Unit.|  "
  Salida.Row = 0
  Salida.Cols = 6
  Salida.ColWidth(0) = 1200
  Salida.ColWidth(1) = 2600
  Salida.ColWidth(2) = 890
  Salida.ColWidth(3) = 1200
  Salida.ColWidth(4) = 1200
  Salida.ColWidth(5) = 1200
End Sub

Private Sub limpiaTexto()
   Text1 = ""
   Text11 = ""
   'TxtNroGuia = ""
   Label3 = ""
   nroguia = ""
End Sub

Public Sub grabastk()
   Dim RSQL As String
   Dim cadena As String
   Dim auxdisp As Double
   
   RSQL = "select * from stkart where STCODIGO= '" & Text1 & "' and  '" & VGAlma & "'"
   Adoreg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
   If Adoreg1.RecordCount = 0 Then
       Adoreg1.AddNew
       Adoreg1("STALMA") = VGAlma   '"01"
       Adoreg1("STCODIGO") = Text1   'acodigo
       Adoreg1("STKFECULT") = DTPicker1
   Else
       auxdisp = Adoreg1("STSKDIS")
   End If
   If Adoreg1("STKPREPRO") <> 0 Then  'no se registrado algun precio
        Adoreg1("STKPREPRO") = (promedio * CANTIDAD + auxdisp * Adoreg1("STKPREPRO")) / (CANTIDAD + auxdisp)
        Adoreg1("STSKDIS") = CANTIDAD + auxdisp
   Else
      Adoreg1("STKPREPRO") = promedio
      Adoreg1("STKPREULT") = promedio
      Adoreg1("STSKDIS") = CANTIDAD
   End If
   Adoreg1.UpdateBatch
   Adoreg1.Requery
   Set Adoreg1 = Nothing
End Sub
Private Sub grabacabecera()
'   Set Adoreg1 = New adodb.Recordset
'  Dim rsql As String
'  rsql = "select * from movalmcab"
'  Adoreg1.Open rsql, cconexcom, adOpenDynamic, adLockOptimistic
'
'         Adoreg1.AddNew
'         Adoreg1("canumdoc") = nroguia
'         Adoreg1("cafecha") = DTPicker1
'         Adoreg1("codigo") = Salida.TextMatrix(i, 0)
'         Adoreg1("descri") = Salida.TextMatrix(i, 1)
'         Adoreg1("unidad") = Salida.TextMatrix(i, 2)
'         Adoreg1("cantidad") = Salida.TextMatrix(i, 3)
'         Adoreg1("precio") = Salida.TextMatrix(i, 4)
'        adoreg1.UpdateBatch
End Sub

Public Function existe_codigo_mabel(cCod As String) As Boolean  ' Codigo del ARTICULO
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
     existe_codigo_mabel = False
     Exit Function
End If
csql = "Select ACODIGO ,CANTIDAD   from guiamabel where CODIGO = '" & Trim(cCod) & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, cConexCom, adOpenStatic
If cSelC.RecordCount > 0 Then
    Text1 = cSelC(0)
    cantidadxCodigo = cSelC(1)
    existe_codigo_mabel = False: cSelC.Close
    Exit Function
End If
 existe_codigo_mabel = True: cSelC.Close
End Function

Private Sub TxtNroGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Trim(TxtNroGuia) = "" Or Len(TxtNroGuia) <> 10 Then
        MsgBox "Ingrese el numero de Guia correcto", vbInformation, "Informacion"
        TxtNroGuia.SetFocus
    Else
        Text11_KeyPress (13)
    End If
End If
End Sub
