VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmRegistrarIASA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar Factura de IASA"
   ClientHeight    =   3075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "FrmRegistrarIASA.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3075
   ScaleWidth      =   5535
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   3360
      Picture         =   "FrmRegistrarIASA.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2160
      Width           =   775
   End
   Begin VB.CommandButton CmdRegistrar 
      Caption         =   "&Enviar"
      Height          =   675
      Left            =   1680
      Picture         =   "FrmRegistrarIASA.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2160
      Width           =   775
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox Text11 
         Height          =   285
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   3135
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
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox TxtNroGuia 
         Height          =   285
         Left            =   1080
         MaxLength       =   10
         TabIndex        =   1
         Top             =   840
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   503
         _Version        =   393216
         Format          =   53477377
         CurrentDate     =   36686
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Ubicacion"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Nro. Guia"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmRegistrarIASA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Adoreg1 As ADODB.Recordset
Dim AdoReg2 As ADODB.Recordset
Dim adoreg3 As ADODB.Recordset
Dim adoreg4 As ADODB.Recordset
Dim rs As Recordset
Dim contador  As Integer
Dim nrguiamabel As String
Dim CANTIDAD As Double

Dim rsql As String
Dim rutalectura, NombreTabla As String
Dim numfac As String
Dim promedio As Double
Dim Cod As String

Private Sub CmdRegistrar_Click()
   Dim rsql As String
   Dim NumDoc As String
   On Error GoTo Err
   'colocar un wait
   
   If InStr(Trim(Text11), ".dbf") Then
        Text11 = Trim(Text11)
        NombreTabla = Mid(Text11, InStrRev(Text11, "\") + 1)
        NombreTabla = Left(NombreTabla, InStrRev(NombreTabla, ".") - 1)
        rutalectura = Left(Text11, InStr(Text11, ":"))
   End If
   
   If Trim(TxtNroGuia) = "" Or Len(TxtNroGuia) <> 10 Then
      MsgBox "Ingrese el numero de la Guia", vbExclamation, "Error"
      Exit Sub
   End If
    Screen.MousePointer = 11
    rsql = "select nroguia from guiamabel where giasa= '" & TxtNroGuia & "' "
    Set adoreg4 = New ADODB.Recordset
    adoreg4.Open rsql, VGcnx, adOpenStatic
    If Not adoreg4.EOF Then
          NumDoc = adoreg4(0)
          Set adoreg4 = Nothing
           rsql = "select * from movalmcab where  catd='NI' and canumdoc='" & NumDoc & "'"
          Set adoreg4 = New ADODB.Recordset
          adoreg4.Open rsql, VGcnx, adOpenStatic
          If Not adoreg4.EOF Then
                   MsgBox "Ya se registro el documento", vbExclamation, "Aviso"
                   Screen.MousePointer = 1
                  Set adoreg4 = Nothing
                  Exit Sub
          End If
    Else
         Set adoreg4 = Nothing
    End If
   
    grabaguiamabel
    If Trim(numfac) <> "" Then
          grabastk_cabdet
          MsgBox "Transferencia concluída", vbInformation, "Aviso"
          CmdRegistrar.Enabled = False
    End If
    Screen.MousePointer = 1
    Exit Sub
Err:
  If Err.Number = 3265 Then
         MsgBox "El archivo  " & NombreTabla & " no coresponde al formato de una factura,verifique", vbInformation, "Aviso"
  ElseIf -2147467259 = Err.Number Then
        MsgBox "Ya se registro el documento", vbExclamation, "Aviso"
  Else
        MsgBox Err.Description, vbInformation, "Aviso"
  End If
  Screen.MousePointer = 1
End Sub

Private Sub CmdSalir_Click()
   Unload Me
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

Private Sub DTPicker1_Change()
        VGTipCamb = DevolverTCambio(DTPicker1.Value)
End Sub

Private Sub Form_Load()
   central Me
   CmdRegistrar.Enabled = False
   DTPicker1 = Date
   VGTipCamb = DevolverTCambio(DTPicker1.Value)
End Sub

Private Sub grabaguiamabel()
'realiza una actualizacion de la guia
 Dim Crsql  As String
' Dim db  As Database
 Dim afectado As Long
'        Set db = Workspaces(0).OpenDatabase(rutalectura, False, False, "FoxPro 2.6")        '     "C:\WINDOWS
'        Set RS = db.OpenRecordset(nombretabla, dbOpenDynaset)
        Set rs = VGcnx.Execute("select * from " & NombreTabla)
        If rs.EOF Then
              MsgBox "No hay registro  en la guia", vbCritical
              End
        End If
        numfac = ""
        rs.MoveFirst
        Set Adoreg1 = New ADODB.Recordset
        rsql = "select * from guiamabel"
        Adoreg1.Open rsql, VGcnx, adOpenDynamic, adLockOptimistic
          If TxtNroGuia <> rs.Fields("refdocum") Then
            MsgBox "El nro de guia ingresado no coincide  con" & Chr(13) & "el nro de guia de remision registrado,verifique", vbExclamation, "Aviso"
            Set Adoreg1 = Nothing
             Exit Sub
        End If
        If IsNull(rs.Fields("refdocum")) Then
            MsgBox "La Factura no está asociada a ninguna Guia", vbInformation, "Información"
            Exit Sub
        Else
            nrguiamabel = numeroguia(rs.Fields("refdocum"))
        End If
        
        If Trim(nrguiamabel) = "" Then
             MsgBox "No se ha registrado la guia de remisión ", vbExclamation, "Aviso"
             Set Adoreg1 = Nothing
             Exit Sub
        End If
        numfac = rs.Fields("factura")
        While Not rs.EOF
               CANTIDAD = rs.Fields("cantidad")
               promedio = rs.Fields("precio")
               Crsql = "update guiamabel set  cantidad= " & CANTIDAD & " , precio = " & promedio & " where nroguia='" & nrguiamabel & "' and codigo ='" & rs.Fields("codigo") & "'"
               VGcnx.Execute Crsql, afectado
               If afectado = 0 Then
                  MsgBox "El codigo: " & rs.Fields("codigo") & " no se encuentra en la guia de remisión" & Chr(13) & " Verifique  la Guia de Remisión con la Factura ", vbInformation, "Aviso"
               End If
               rs.MoveNext
        Wend
        Set Adoreg1 = Nothing
End Sub

Private Sub grabacabecera()
  'Desea grabar el registro
   ' numfac = Format(nombretabla, String(10, "0"))
    Set adoreg3 = New ADODB.Recordset
        rsql = "select * from movalmcab"
        adoreg3.Open rsql, VGcnx, adOpenDynamic, adLockOptimistic
        adoreg3.AddNew
        adoreg3.Fields("CAALMA") = VGAlma     '"01"
        adoreg3.Fields("CANUMDOC") = nrguiamabel
        adoreg3.Fields("CATIPMOV") = "I"
        adoreg3.Fields("CATD") = "NI"
        adoreg3.Fields("CAHORA") = Format(Time, "hh:mm:ss")
        adoreg3.Fields("CAFECDOC") = DTPicker1.Value
        adoreg3.Fields("CARFTDOC") = "FT"
        adoreg3.Fields("CARFNDOC") = numfac
        adoreg3.Fields("CACODMOV") = "CL"
        adoreg3.Fields("CACODPRO") = VGIASA           'txtIASA
        adoreg3.Fields("CANOMPRO") = "IASA"              'LTrim(Label13.Caption)
       adoreg3.Fields("CAUSUARI") = UCase(VGUsua)
       adoreg3.Fields("CACODMON") = "MN"               'VGCodMon
       adoreg3.Fields("CASITGUI") = "V"
       'adoreg3.Fields("CASITUA") = "V"
       adoreg3.Fields("CAESTIMP") = "V"
       adoreg3.Update
       adoreg3.Requery
   Set adoreg3 = Nothing

End Sub

Private Sub grabadetalle()
        Set AdoReg2 = New ADODB.Recordset
        rsql = "select * from movalmdet where  dealma = '" & VGAlma & "'  and denumdoc = '" & nrguiamabel & "' and decodigo = '" & Cod & "'"
        AdoReg2.Open rsql, VGcnx, adOpenDynamic, adLockOptimistic
        '(promedio * cantidad + auxdisp * adoreg3("STKPREPRO")) / (cantidad + auxdisp)
        If AdoReg2.RecordCount = 0 Then
                AdoReg2.AddNew
                AdoReg2.Fields("DEALMA") = VGAlma
                AdoReg2.Fields("DETD") = "NI"
                AdoReg2.Fields("DENUMDOC") = nrguiamabel
                AdoReg2.Fields("DEITEM") = contador
                AdoReg2.Fields("DECODIGO") = Cod
                AdoReg2.Fields("DEDESCRI") = buscadescri(Cod)
                AdoReg2.Fields("DECANTID") = CANTIDAD
                AdoReg2.Fields("DEPRECIO") = promedio
        Else
                AdoReg2.Fields("DEPRECIO") = (promedio * CANTIDAD + AdoReg2.Fields("DECANTID") * AdoReg2.Fields("DEPRECIO")) / (CANTIDAD + AdoReg2.Fields("decantid"))
                AdoReg2.Fields("DECANTID") = AdoReg2.Fields("DECANTID") + CANTIDAD
        End If
        AdoReg2.Update
        AdoReg2.Requery
        Set AdoReg2 = Nothing
End Sub

Public Sub grabastk()
   Dim rsql As String
   Dim CADENA As String
   Dim auxdisp As Double
   
   Set adoreg3 = New ADODB.Recordset
   rsql = "select * from stkart where STCODIGO= '" & Cod & "' and  '" & VGAlma & "'"
   adoreg3.Open rsql, VGcnx, adOpenDynamic, adLockOptimistic
   If adoreg3.RecordCount = 0 Then
       adoreg3.AddNew
       adoreg3("STALMA") = VGAlma   '"01"
       adoreg3("STCODIGO") = Cod  'acodigo
       adoreg3("STKFECULT") = DTPicker1
   Else
       auxdisp = adoreg3("STSKDIS")
   End If
     If adoreg3("STKPREPRO") <> 0 And (CANTIDAD + auxdisp) <> 0 Then  'no se registrado algun precio
        adoreg3("STKPREPRO") = (promedio * CANTIDAD + auxdisp * adoreg3("STKPREPRO")) / (CANTIDAD + auxdisp)
        adoreg3("STSKDIS") = CANTIDAD + auxdisp
     Else
        adoreg3("STKPREPRO") = Round(promedio, 6)
        adoreg3("STKPREULT") = promedio
        adoreg3("STSKDIS") = CANTIDAD
     End If
   adoreg3.UpdateBatch
   adoreg3.Requery
   ValMes
   Set adoreg3 = Nothing
End Sub

Public Function buscadescri(cCod As String) As String  ' Codigo del ARTICULO
Dim cSelC As ADODB.Recordset, csql As String
If Trim(cCod) = "" Then
    buscadescri = " "
    Exit Function
End If
csql = "Select   ADESCRI from MaeART where ACODIGO = '" & Trim(cCod) & "'"
Set cSelC = New ADODB.Recordset
cSelC.Open csql, VGcnx, adOpenStatic
If cSelC.RecordCount > 0 Then
    buscadescri = cSelC(0): cSelC.Close
    Exit Function
End If
     buscadescri = " ": cSelC.Close
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        VGTipCamb = DevolverTCambio(VG_FecTrab)
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And InStr(Trim(Text11), ".dbf") Then
  Text11 = Trim(Text11)
  NombreTabla = Mid(Text11, InStrRev(Text11, "\") + 1)
  NombreTabla = Left(NombreTabla, InStrRev(NombreTabla, ".") - 1)
  rutalectura = Left(Text11, InStr(Text11, ":"))
  TxtNroGuia.SetFocus
End If
End Sub

Public Function numeroguia(cCod As String) As String
Dim cSelC As ADODB.Recordset, csql As String
    If Trim(cCod) = "" Then
        numeroguia = " "
        Exit Function
    End If
    csql = "Select   nroguia from guiamabel where giasa = '" & Trim(cCod) & "'"
    Set cSelC = New ADODB.Recordset
    cSelC.Open csql, VGcnx, adOpenStatic
    If cSelC.RecordCount > 0 Then
       numeroguia = cSelC(0): cSelC.Close
        Exit Function
    End If
         numeroguia = " ": cSelC.Close
End Function

Private Sub grabastk_cabdet()
Dim monto As Double, nCanref As Double
Dim Unidad As String, nCanConv As Double

Set Adoreg1 = New ADODB.Recordset
   rsql = "select * from  guiamabel where nroguia= '" & nrguiamabel & " '   order by acodigo"
   Adoreg1.Open rsql, VGcnx, adOpenDynamic, adLockOptimistic
   If Adoreg1.RecordCount = 0 Then
      MsgBox "no existe  datos en la base de datos", vbInformation, "Aviso"
      Exit Sub
   End If
   grabacabecera
   contador = 0
   CANTIDAD = 0
   monto = 0
   promedio = 0
   codigoaux = Adoreg1("acodigo")
   unidadaux = Adoreg1("unidad")
   nCanConv = 1
   While Not Adoreg1.EOF
       If codigoaux <> Adoreg1("acodigo") Or unidadaux <> Adoreg1("unidad") Then
         contador = contador + 1
          If CANTIDAD <> 0 Then
              promedio = Round(monto / (CANTIDAD * nCanConv), 8)
         End If
         Cod = codigoaux
         grabastk
         grabadetalle
         codigoaux = Adoreg1("acodigo")
         unidadaux = Adoreg1("unidad")
         CANTIDAD = 0
         monto = 0
         promedio = 0
         nCanConv = 1
       End If
       Unidad = "": nCanref = 0
       Unidad = Devolver_Dato(1, Adoreg1("acodigo"), "MaeArt", "acodigo", False, "Aunidad")
       nCanConv = Val(Devolver_Dato(1, Adoreg1("unidad"), "TABEQUI", "EQUNIPRI", False, "EQCANTEQUI", Unidad, "EQUNIEQUI"))
       nCanref = Val(Devolver_Dato(1, Adoreg1("unidad"), "TABEQUI", "EQUNIPRI", False, "EQCANTEQUI", Unidad, "EQUNIEQUI"))
       If nCanref = 0 Then nCanref = 1
       If nCanConv = 0 Then nCanConv = 1
       nCanref = Adoreg1("cantidad") * nCanref
       CANTIDAD = nCanref + CANTIDAD
       monto = monto + nCanref * Adoreg1("precio")
       Adoreg1.MoveNext
   Wend
       contador = contador + 1
       If CANTIDAD <> 0 Then
              'Se hace la conversión en Base a la Unidad de Medida en la que viene, si el precio es Doc, la cantidad  tiene que ser Docenas
              promedio = Round(monto / (CANTIDAD * nCanConv), 8)
       End If
       Cod = codigoaux
       grabastk
       grabadetalle
End Sub

Private Sub TxtNroGuia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If Len(TxtNroGuia) <> 10 Then
        MsgBox "Ingrese el Nro. de guia correcto ", vbExclamation, "Error"
   Else
        If InStr(Trim(Text11), ".dbf") Then
          Text11 = Trim(Text11)
          NombreTabla = Mid(Text11, InStrRev(Text11, "\") + 1)
          NombreTabla = Left(NombreTabla, InStrRev(NombreTabla, ".") - 1)
          rutalectura = Left(Text11, InStr(Text11, ":"))
          CmdRegistrar.Enabled = True
        End If
   End If
End If
End Sub

Private Sub ValMes()
 
  Dim rsql As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  On Error GoTo Err
   mespro = Year(DTPicker1) & Format(Month(DTPicker1), "00")
   'cadena = MsFlexGrid1.TextMatrix(contador, 0) 'codigo del art
   rsql = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & Cod & "'"  '
   Set adoreg4 = New ADODB.Recordset
   adoreg4.Open rsql, VGcnx, adOpenDynamic, adLockOptimistic
   If Not adoreg4.EOF Then 'existe
       Cantent = adoreg4(0) + CANTIDAD
       uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & VGAlma & "'  and  SMCODIGO ='" & Cod & "' AND SMMESPRO ='" & mespro & "' "
   Else
       Cantent = CANTIDAD
       Cantsal = 0
       uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & VGAlma & "','" & Cod & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
   End If
   VGcnx.Execute uSql
   adoreg4.Close
  Exit Sub
Err:
   MsgBox Err.Description, vbExclamation, "Error !!!"
End Sub
