VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmTraIng 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ingreso a Almacén"
   ClientHeight    =   5595
   ClientLeft      =   120
   ClientTop       =   4320
   ClientWidth     =   9225
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTraIng.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9225
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   675
      Left            =   4536
      Picture         =   "frmTraIng.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4755
      Width           =   775
   End
   Begin VB.TextBox txtCol 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4608
      MultiLine       =   -1  'True
      TabIndex        =   30
      Top             =   4860
      Visible         =   0   'False
      Width           =   552
   End
   Begin VB.CommandButton cmdGra 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   675
      Left            =   3528
      Picture         =   "frmTraIng.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4752
      Width           =   775
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Flex1 
      Height          =   1956
      Left            =   72
      TabIndex        =   5
      Top             =   2664
      Width           =   9084
      _ExtentX        =   16007
      _ExtentY        =   3440
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      RowHeightMin    =   240
      BackColorSel    =   -2147483643
      ForeColorSel    =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      ScrollBars      =   2
      Appearance      =   0
      FormatString    =   "^Cód Art|Descripción|Uni.|>Cantidad|>Saldo|>Cant.Recibida"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   7.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame fraCabec 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   72
      TabIndex        =   17
      Top             =   0
      Width           =   9108
      Begin VB.TextBox txtNum 
         Height          =   285
         Left            =   2040
         MaxLength       =   13
         TabIndex        =   0
         Top             =   210
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Número Importación:"
         Height          =   195
         Left            =   375
         TabIndex        =   22
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.Frame fraDatos 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2145
      Left            =   72
      TabIndex        =   8
      Top             =   405
      Width           =   9096
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   270
         Left            =   1455
         TabIndex        =   31
         Top             =   1215
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   476
         _Version        =   393216
         Format          =   24903681
         CurrentDate     =   36942
      End
      Begin VB.TextBox txtTF 
         Height          =   285
         Left            =   4920
         TabIndex        =   1
         Top             =   1200
         Width           =   375
      End
      Begin VB.TextBox txtNTF 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5400
         TabIndex        =   2
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtTM 
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Height          =   285
         Left            =   6240
         MaxLength       =   2
         TabIndex        =   4
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtAlm 
         Height          =   285
         Left            =   1470
         MaxLength       =   3
         TabIndex        =   3
         Top             =   1560
         Width           =   315
      End
      Begin VB.Label lblTM 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6720
         TabIndex        =   29
         Top             =   1560
         Width           =   1830
      End
      Begin VB.Label lblTF 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6720
         TabIndex        =   28
         Top             =   1200
         Width           =   1830
      End
      Begin VB.Label lblEst 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6480
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label lblEsta 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   6960
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
         Width           =   1590
      End
      Begin VB.Label lblEnt 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4080
         TabIndex        =   25
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblEmi 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1455
         TabIndex        =   24
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblPro 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1455
         TabIndex        =   23
         Top             =   360
         Width           =   1110
      End
      Begin VB.Label lblAlm 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   1920
         TabIndex        =   21
         Top             =   1560
         Width           =   2535
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Movim. :"
         Height          =   195
         Left            =   5160
         TabIndex        =   20
         Top             =   1575
         Width           =   960
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Estado  :"
         Height          =   195
         Left            =   5640
         TabIndex        =   19
         Top             =   735
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Entrega   :"
         Height          =   195
         Left            =   3240
         TabIndex        =   18
         Top             =   735
         Width           =   735
      End
      Begin VB.Label lblRuc 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   7440
         TabIndex        =   16
         Top             =   360
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "R.U.C.  :"
         Height          =   195
         Left            =   6720
         TabIndex        =   15
         Top             =   375
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor     :"
         Height          =   195
         Left            =   375
         TabIndex        =   14
         Top             =   380
         Width           =   1005
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Emisión         :"
         Height          =   195
         Left            =   375
         TabIndex        =   13
         Top             =   740
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha           :"
         Height          =   195
         Left            =   375
         TabIndex        =   12
         Top             =   1215
         Width           =   990
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Almacén       :"
         Height          =   195
         Left            =   375
         TabIndex        =   11
         Top             =   1575
         Width           =   975
      End
      Begin VB.Label lblProv 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   2640
         TabIndex        =   10
         Top             =   360
         Width           =   2970
      End
      Begin VB.Label lblCen 
         AutoSize        =   -1  'True
         Caption         =   "Tip./Fact.    :"
         Height          =   195
         Left            =   3840
         TabIndex        =   9
         Top             =   1215
         Width           =   930
      End
   End
End
Attribute VB_Name = "frmTraIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Adodc1 As ADODB.Recordset
Dim Adodc2 As ADODB.Recordset
Dim nTra As Integer
Dim Mensaje As String
Public nro_serie As Integer

Private Sub cmdGra_Click()
    Dim SQLc As String
    Dim SQLd As String
    Dim i As Integer           'Contiene la posicion del flex
    Dim consecutivo As Integer 'Contiene el numero de item
    Dim totalserie As Integer  'Contiene cuantos items se ingresa por producto
    Dim TipoIng As Integer
    Dim vNI As Integer
    Dim vNF As Double, vNC As Double
    Dim vNP As Double, vNP1 As Double
    Dim criterio As String
    Dim csql As String
    On Error GoTo GrabErr
    Dim xSaldo As Double
    xSaldo = 0
    txtTF = Trim(txtTF)
    If Trim(txtTF) = "" Then
'        mensaje = "Debe especificar Tipo de Documento"
'        MsgBox mensaje, vbExclamation, "Mensaje"
'        txtTF.SetFocus
'        Exit Sub
    Else
        If Not Existe(1, txtTF, "tipo_docu", "tdo_tipdoc", False) Then
            Mensaje = "No existe el Tipo de Documento ingresado"
            MsgBox Mensaje, vbExclamation, "Error"
            txtTF.SetFocus
            Exit Sub
        Else
            If lblTF = "" Then
                lblTF = Devolver_Dato(1, txtTF, "tipo_docu", "tdo_tipdoc", False, _
                    "tdo_descri")
                txtNTF.Enabled = True
            End If
        End If
    End If
    
    txtNTF = Trim(txtNTF)
    If txtNTF = "" Then
'        mensaje = "Debe especificar el Número de Documento"
'        MsgBox mensaje, vbExclamation, "Mensaje"
'        txtNTF.SetFocus
'        Exit Sub
    End If

    txtAlm = Trim(txtAlm)
    If txtAlm = "" Then
        Mensaje = "Debe especificar Código de Almacen"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        txtAlm.SetFocus
        Exit Sub
    Else
        If Not Existe(1, txtAlm, "tabalm", "taalma", False) Then
            Mensaje = "El Código de Almacén ingresado no existe"
            MsgBox Mensaje, vbExclamation, "Error"
            txtAlm.SetFocus
            Exit Sub
        Else
            If lblAlm = "" Then lblAlm = Devolver_Dato(1, txtAlm, "tabalm", "taalma", _
                False, "tadescri")
        End If
    End If

    txtTM = Trim(txtTM)
    If txtTM = "" Then
        Mensaje = "Debe especificar Tipo de Movimiento"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        txtTM.SetFocus
        Exit Sub
    Else
        If Not Existe(1, txtTM, "tabtransa", "tt_codmov", False) Then
            Mensaje = "El Tipo de Movimiento ingresado no existe"
            MsgBox Mensaje, vbExclamation, "Error"
            txtTM.SetFocus
        Else
            If lblTM = "" Then lblTM = Devolver_Dato(1, txtTM, "tabtransa", "tt_codmov", _
                False, "tt_descri")
        End If
    End If

    TipoIng = Ingreso_Realizado
    If TipoIng = 0 Then
        Mensaje = "No se puede grabar." & vbCrLf & "No se ha recepcionado ningún artículo"
        MsgBox Mensaje, vbExclamation, "Mensaje"
        Flex1.SetFocus
        Exit Sub
    End If

    Mensaje = "¿Desea guardar los cambios realizados?"
    If MsgBox(Mensaje, vbQuestion + vbYesNo, "Mensaje") = vbYes Then
        'Obtengo el consecutivo
        vNI = Devolver_Dato(1, txtAlm, "tabalm", "taalma", False, "tanument")
        vNI = vNI + 1
        
        nTra = 1
        
        cConexCom.BeginTrans
                    
        
        SQLd = ""
        SQLd = "INSERT INTO movalmcab (caalma,catd,canumdoc,cafecdoc,catipmov,cacodmov," & _
            "casitua,carftdoc,carfndoc,cacodpro,cafecact,cahora,causuari,cacodmon," & _
            "canroimp,casitgui) VALUES ('" & txtAlm & "','NI','" & Format(vNI, "0000000000") & _
            "'," & DateSQL(DTPicker1) & ",'I','" & txtTM & "','M','" & _
            txtTF & "','" & txtNTF & "','" & lblPro & "'," & DateSQL(Date) & _
            ",'" & Format(Time, "hh:mm:ss") & "','" & VGUsuario & "','ME','" & _
            txtNum & "','V')"
        cConexCom.Execute SQLd
        consecutivo = 1
        
        MsgBox "El almacen de ingreso es " & txtAlm, vbInformation, "Aviso"
        For i = 1 To Flex1.Rows - 1
            If Val(Flex1.TextMatrix(i, 5)) > 0 Then
                vNP = Devolver_Dato(1, txtNum, "impord", "cnumero", False, "npreneto", _
                    Flex1.TextMatrix(i, 0), "ccodartic")
                vNP1 = vNP * Val(Flex1.TextMatrix(i, 5))
                SQLd = ""
                
                Set Adodc1 = New ADODB.Recordset
                csql = "Select AFLOTE,AFSERIE  from MAEART  where ACODIGO ='" & Flex1.TextMatrix(i, 0) & "' "
                Adodc1.Open csql, cConexCom, adOpenStatic, adLockReadOnly
                If Trim(Adodc1(1)) = "S" Then 'RMM*****PARA EL CASO DE series
                    totalserie = 0
                    Set Adodc2 = New ADODB.Recordset
                    csql = "Select SERIE  from ART_SERIE  where ACODIGO ='" & Flex1.TextMatrix(i, 0) & "' AND ALMA ='" & txtAlm & "' "
                    Adodc2.Open csql, cConexCom, adOpenStatic, adLockReadOnly
                    
                    If Adodc2.EOF Then
                       Adodc2.Close
                       Adodc1.Close
                       MsgBox "No se registro ninguna serie ,vuelva a realizar la transaccion", vbInformation, "Aviso"
                       cConexCom.RollbackTrans
                       cConexCom.Execute " delete from movalmcab where canumdoc = '" & Format(vNI, "0000000000") & "' and caALMA ='" & txtAlm & "' and catd ='NI'"
                       cConexCom.Execute " delete from movalmdet where denumdoc = '" & Format(vNI, "0000000000") & "' and deALMA ='" & txtAlm & " 'and detd ='NI'"
                       Exit Sub
                    End If
                    
                    If Int(Adodc2.RecordCount) <> Int(Val(Flex1.TextMatrix(i, 5))) Then
                       Adodc2.Close
                       Adodc1.Close
                       MsgBox "La cantidad  de series ingresado no esta completo, vuelva a realizar la transaccion", vbInformation, "Aviso"
                       cConexCom.RollbackTrans
                       cConexCom.Execute " delete from movalmcab where canumdoc = '" & Format(vNI, "0000000000") & "' and caALMA ='" & txtAlm & "' and catd ='NI'"
                       cConexCom.Execute " delete from movalmdet where denumdoc = '" & Format(vNI, "0000000000") & "' and deALMA ='" & txtAlm & "' and detd ='NI'"
                       Exit Sub
                    End If
                    
                    vNP = vNP '/ Val(Flex1.TextMatrix(i, 5))
                    Adodc2.MoveFirst
                    While totalserie < Val(Flex1.TextMatrix(i, 5))
                        SQLd = "INSERT INTO movalmdet (dealma,detd,denumdoc,deitem,decodigo," & _
                            "decantid,deprecio,deestado,decodmov,devaltot," & _
                            "decodmon,deserie,deitemi) VALUES ('" & txtAlm & "','NI','" & Format(vNI, "0000000000") & _
                            "'," & consecutivo & ",'" & Flex1.TextMatrix(i, 0) & "',1," & vNP & ",'" & _
                            "V','" & txtTM & "'," & _
                            vNP1 & ",'ME','" & Adodc2("serie") & "'," & i & ")"
                        cConexCom.Execute SQLd
                        Adodc2.MoveNext
                        consecutivo = consecutivo + 1
                        totalserie = totalserie + 1
                    Wend
                Else
                     If Trim(Adodc1(0)) = "S" Then 'RMM*****PARA EL CASO DE LOTES
'RMM***************************************************************************************************************
                         Set Adodc2 = New ADODB.Recordset
                         csql = "Select LOTE,CANTID  from ART_LOTE  where ACODIGO ='" & Flex1.TextMatrix(i, 0) & "' AND ALMA ='" & txtAlm & "'"
                         Adodc2.Open csql, cConexCom, adOpenStatic, adLockReadOnly
                         
                         If Adodc2.EOF Then
                            Adodc2.Close
                            Adodc1.Close
                            MsgBox "No se registro ningun Lote ,vuelva a realizar la transaccion", vbInformation, "Aviso"
                            cConexCom.RollbackTrans
                            cConexCom.Execute " delete from movalmcab where canumdoc = '" & Format(vNI, "0000000000") & "' and caALMA ='" & txtAlm & "' and catd ='NI'"
                            cConexCom.Execute " delete from movalmdet where denumdoc = '" & Format(vNI, "0000000000") & "' and deALMA ='" & txtAlm & " 'and detd ='NI'"
                            Exit Sub
                         End If
                        
                         Adodc2.MoveFirst
                         
                         Do While Not Adodc2.EOF
                             SQLd = "INSERT INTO movalmdet (dealma,detd,denumdoc,deitem,decodigo," & _
                                "decantid,deprecio,deestado,decodmov,devaltot," & _
                                "decodmon,deserie,deitemi) VALUES ('" & txtAlm & "','NI','" & Format(vNI, "0000000000") & _
                                "'," & consecutivo & ",'" & Flex1.TextMatrix(i, 0) & "'," & Adodc2("CANTID") & "," & vNP & ",'" & _
                                "V','" & txtTM & "'," & _
                                vNP1 & ",'ME','" & Adodc2("LOTE") & "'," & i & ")"
                             cConexCom.Execute SQLd
                             consecutivo = consecutivo + 1
                             Adodc2.MoveNext
                         Loop
'***************************************************************************************************************
                     Else  'RMM*****PARA EL CASO DE articulo normal
                     
                         SQLd = "INSERT INTO movalmdet(dealma,detd,denumdoc,deitem,decodigo," & _
                             "decantid,deprecio,deestado,decodmov,devaltot," & _
                             "decodmon,deitemi) VALUES ('" & txtAlm & "','NI','" & Format(vNI, "0000000000") & _
                             "'," & consecutivo & ",'" & Flex1.TextMatrix(i, 0) & "'," & Val(Flex1.TextMatrix(i, 5)) & "," & vNP & ",'" & _
                             "V','" & txtTM & "'," & _
                             vNP1 & ",'ME'," & i & ")"
                         cConexCom.Execute SQLd
                         consecutivo = consecutivo + 1
                     End If
                    
                End If
                Adodc1.Close
            End If
        Next
        
        SQLd = ""
        For i = 1 To Flex1.Rows - 1
            If Val(Flex1.TextMatrix(i, 5)) > 0 Then
                xSaldo = xSaldo + Val(Flex1.TextMatrix(i, 4)) - Val(Flex1.TextMatrix(i, 5))
                SQLd = "UPDATE impord SET ncantentr= ncantentr+" & Val(Flex1.TextMatrix(i, 5)) & _
                    ",ncansaldo=" & Val(Flex1.TextMatrix(i, 4)) - _
                    Val(Flex1.TextMatrix(i, 5)) & " WHERE cnumero='" & txtNum & _
                    "' AND ccodartic='" & Flex1.TextMatrix(i, 0) & "'"
                cConexCom.Execute SQLd
            End If
        Next
        
        'actualiza el stock
        For i = 1 To Flex1.Rows - 1
            If Existe(1, Flex1.TextMatrix(i, 0), "stkart", "stcodigo", False, txtAlm, _
                "stalma") Then
                SQLd = "UPDATE stkart SET stskdis=stskdis+" & _
                    Val(Flex1.TextMatrix(i, 5)) & " WHERE stalma='" & txtAlm & _
                    "'AND stcodigo='" & Flex1.TextMatrix(i, 0) & "'"
            Else
                SQLd = "INSERT INTO stkart (stalma,stcodigo,stskdis) VALUES ('" & txtAlm & _
                    "','" & Flex1.TextMatrix(i, 0) & "'," & Val(Flex1.TextMatrix(i, 5)) & ")"
            End If
            cConexCom.Execute SQLd
        Next
        'actualiza el consecutivo
         SQLc = "UPDATE tabalm SET tanument=" & vNI & " WHERE taalma='" & txtAlm & "'"
         cConexCom.Execute SQLc
         'adodc1.Requery
         Dim cAlma As String
         cAlma = txtAlm
         Mensaje = "Se ingresó la mercadería." & vbCrLf & vbCrLf & "Nota de Ingreso : " & _
             Format(vNI, "0000000000") & vbCrLf & "Almacén : " & txtAlm & vbCrLf & _
             "Tipo de Movimiento : " & txtTM
         MsgBox Mensaje, vbInformation, "Ingreso"
         
    
         '*RMM ACTUALIZA SERIE************************************************************************
         '********************************************************************************************
         criterio = "select * from art_serie where alma = '" & cAlma & "'"
         Set Adodc1 = New ADODB.Recordset
         Adodc1.Open criterio, cConexCom, adOpenDynamic, adLockBatchOptimistic
         If Not Adodc1.EOF Then
             Set Adodc2 = New ADODB.Recordset
             Adodc2.Open "select * from stkseri", cConexCom, adOpenDynamic, adLockBatchOptimistic
             With Adodc2
                 While Not Adodc1.EOF
                      .AddNew
                      .Fields("stsalma") = cAlma
                      .Fields("stscodigo") = Adodc1("acodigo")        ' form.text1  qwue tiene el codigo
                      .Fields("stsserie") = Adodc1("serie")
                      .Fields("stsskdis") = 1
                      .UpdateBatch
                      Adodc1.MoveNext
                 Wend
             End With
             Adodc2.Close
         End If
         Adodc1.Close
        '*RMM ACTUALIZA LOTE************************************************************************
        'EL LOTE QUE SELECCIONO YA EXISTE EN STKLOTE
        '********************************************************************************************
        criterio = "select * from art_lote where alma = '" & cAlma & "'"
        Set Adodc1 = New ADODB.Recordset
        Adodc1.Open criterio, cConexCom, adOpenStatic, adLockReadOnly
        While Not Adodc1.EOF
              cConexCom.Execute "UPDATE STKLOTE SET STSLKDIS=STSLKDIS+" & Adodc1("CANTID") & " WHERE stsalma='" & cAlma & "' AND stscodigo='" & Adodc1("acodigo") & "' AND STSLOTE='" & Adodc1("lote") & "'"
              Adodc1.MoveNext
        Wend
        Adodc1.Close
    
       If xSaldo = 0 Then cConexCom.Execute "update imporc set csituacion='04' where cnumero='" & txtNum & "' and csituacion<>'04'"
       If xSaldo <> 0 Then cConexCom.Execute "update imporc set csituacion='03' where cnumero='" & txtNum & "' and csituacion<>'04'"
           
         cConexCom.CommitTrans
         nTra = 0
         
         txtNum = ""
         txtNum.SetFocus
       
    End If
        
    Exit Sub

GrabErr:
    MsgBox Err.Description
    If nTra = 1 Then cConexCom.RollbackTrans
    Resume
End Sub

Private Sub CmdSalir_Click()
    'Unload frmReferencia
    'Unload frmTraEmi1
    Unload Me
End Sub

Private Sub DTPicker1_Change()
        DTPicker1.Value = UltimoCierreFech(DTPicker1.Value)
        VGTipCamb = DevolverTCambio(DTPicker1.Value)
End Sub

Private Sub DTPicker1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  txtTF.SetFocus
End If
End Sub

Private Sub Flex1_KeyPress(KeyAscii As Integer)
Dim csql As String

    If lblPro = "" Then Exit Sub
    If Flex1.Col = 5 And Val(Flex1.TextMatrix(Flex1.Row, 4)) > 0 Then
        If (KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = 46 Then
            txtCol.FontName = Flex1.CellFontName
            txtCol.FontSize = Flex1.CellFontSize
            txtCol.Width = Flex1.CellWidth
            txtCol.Height = Flex1.CellHeight
            txtCol.Left = Flex1.Left + Flex1.CellLeft
            txtCol.Top = Flex1.Top + Flex1.CellTop
            txtCol.Visible = True
            txtCol = Chr(KeyAscii)
            txtCol.SelStart = 1
            CmdSalir.Cancel = False
           
            txtCol.SetFocus
        End If
        If KeyAscii = 13 Then
        
              If Trim(txtAlm) = "" Then
                 MsgBox "Seleccione primero el Almacen, para poder registrar las Series  ", vbInformation, "Aviso...!"
                 Exit Sub
              Else
                 txtAlm.Enabled = False
              End If
              
              
              formIngSerie.almacen = txtAlm
              VGcod = Flex1.TextMatrix(Flex1.Row, 0)
              nro_serie = Flex1.TextMatrix(Flex1.Row, 5)
              Set Adodc1 = New ADODB.Recordset
              csql = "Select AFLOTE,AFSERIE  from MAEART  where ACODIGO ='" & Flex1.TextMatrix(Flex1.Row, 0) & "' "             '
              Adodc1.Open csql, cConexCom, adOpenDynamic, adLockOptimistic
              If Adodc1(0) = "S" Then
                 If Not CargaLista(Flex1.TextMatrix(Flex1.Row, 0)) Then Call VerLotes
              ElseIf Adodc1(1) = "S" Then
                 formIngSerie.Show 1
              End If
              Adodc1.Close
        End If
    End If
End Sub

Private Sub Form_Load()
    central Me
    
    'Load frmReferencia
    Formato_FlexGrid
    'RMM*******************************************************************
    DTPicker1.Value = UltimoCierreFech(CDate(Format(Now, "dd/MM/yyyy")))
    '*******************************************************************
    VGTipCamb = DevolverTCambio(DTPicker1.Value)
    frmTraIng.txtAlm = VGAlma
End Sub

Sub Limpiar()
    lblPro = "": lblProv = "": lblRuc = ""
    lblEmi = "": lblEnt = "": lblEst = ""
    lblEsta = "":  txtTF = ""
    txtAlm = "": txtTM = ""
    Vacia_FlexGrid
End Sub

Sub Formato_FlexGrid()
    Flex1.FormatString = "Cód Art|Descripción|Uni.|Cantidad|Saldo|Cant.Recibida|Serie/Lote"
    Flex1.ColWidth(0) = 1100
    Flex1.ColWidth(1) = 2900
    Flex1.ColWidth(2) = 550
    Flex1.ColWidth(3) = 1200
    Flex1.ColWidth(4) = 1200
    Flex1.ColWidth(5) = 1200
    Flex1.ColWidth(6) = 500
    
End Sub

Sub Vacia_FlexGrid()
    Dim i As Integer
    
    Do While Flex1.Rows - 1 > 1
        Flex1.RemoveItem 1
    Loop
    
    For i = 0 To 5
        Flex1.TextMatrix(1, i) = ""
    Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        VGTipCamb = DevolverTCambio(VG_FecTrab)
End Sub

Private Sub txtAlm_Change()
    If lblAlm <> "" Then lblAlm = ""
End Sub

Private Sub txtAlm_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT taalma,tadescri FROM tabalm"
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Lista de Almacenes"
    frmReferencia.Show vbModal
    Adodc2.Close
    
    If vGUtil(1) <> "" Then
        txtAlm = vGUtil(1)
        lblAlm = vGUtil(2)
        txtTM.SetFocus
    End If
End Sub

Private Sub txtAlm_GotFocus()
    Enfoque txtAlm
End Sub

Private Sub txtAlm_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtAlm_DblClick
End Sub

Private Sub txtAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtAlm = Trim(txtAlm)
        If txtAlm <> "" Then
            If Not Existe(1, txtAlm, "tabalm", "taalma", False) Then
                Mensaje = "El Código de Almacén ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtAlm.SetFocus
            Else
                lblAlm = Devolver_Dato(1, txtAlm, "tabalm", "taalma", False, "tadescri")
                txtTM.SetFocus
            End If
        Else
            txtTM.SetFocus
        End If
    End If
    'Enteros_Positivos KeyAscii, txtAlm
End Sub

Private Sub txtCol_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Val(Flex1.TextMatrix(Flex1.Row, 4)) < Val(txtCol) Then
            txtCol = Format(txtCol, "0.00")
            txtCol.SelStart = 0
            txtCol.SelLength = Len(txtCol)
        Else
            Flex1.text = Format("0" & txtCol, "0.00")
            CmdSalir.Cancel = True
            txtCol.Visible = False
            Flex1.SetFocus
        End If
    ElseIf KeyAscii = 27 Then
        KeyAscii = 0
        CmdSalir.Cancel = True
        txtCol.Visible = False
        Flex1.SetFocus
    Else
        'Reales_Positivos KeyAscii, txtCol
    End If
End Sub

Private Sub txtCol_LostFocus()
    txtCol.Visible = False
End Sub




Private Sub txtNTF_GotFocus()
    Enfoque txtNTF
End Sub

Private Sub txtNTF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtNTF = Trim(txtNTF)
        Tabula (KeyAscii)
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub Txtnum_Change()
    If lblPro <> "" Then
        Limpiar
        fraDatos.Enabled = False
        cmdGra.Enabled = False
    End If
End Sub

Private Sub txtNum_DblClick()
    Static Adodc1 As ADODB.Recordset
    Dim strsql As String
    On Local Error GoTo ERRAR
    
    Set Adodc1 = New ADODB.Recordset
    
    strsql = "SELECT a.cnumero,a.fentrega,b.est_nombre FROM imporc a,estado_OC b where a.csituacion=b.est_codigo and (a.csituacion='01' or a.csituacion='03')"
'    strsql = "SELECT a.oc_cnumord, a.oc_dfecdoc, b.est_nombre FROM comovc AS a, estado_oc AS b" & _
             " WHERE a.oc_csitord=b.est_codigo"
    Adodc1.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia1.Conectar Adodc1, strsql
    frmReferencia1.lblTit = "Ordenes de Importación"
    frmReferencia1.Inicio
    frmReferencia1.Show vbModal
    Adodc1.Close
    
    If vGUtil(1) <> "" Then
        txtNum = vGUtil(1)
      '  lblSol = vGUtil(2)
      '  txtcen.SetFocus
    End If
Exit Sub
ERRAR:
     If Err.Number = -2147217865 Then
        MsgBox "Usted no Tiene la Interface de Importación ", vbCritical, "Error"
     Else
        MsgBox Err.Description
     End If
End Sub

Private Sub txtNum_GotFocus()
    Enfoque txtNum
End Sub

Private Sub txtNum_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 112 Then txtNum_DblClick
End Sub

Private Sub txtnum_KeyPress(KeyAscii As Integer)
Dim SQL As String
    If KeyAscii = 13 Then
        If txtNum <> "" Then
            'txtNum = Format(txtNum, "0000000000000")
            If Not Existe(1, txtNum, "imporc", "cnumero", False) Then
                Mensaje = "El Número de Orden de Importación ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Mensaje"
                Enfoque txtNum
                txtNum.SetFocus
                Exit Sub
            Else
                Muestra_datos_de_COMOVC
                Rellena_FlexGrid
                txtAlm.Enabled = True
                fraDatos.Enabled = True
                cmdGra.Enabled = True
                Flex1.Col = 5
                If Not ExisteElem(0, cConexCom, "ART_SERIE") Then
                      SQL = " Create Table ART_SERIE (ALMA Text(03),ACODIGO Text(20),SERIE Text(20)) "
                      cConexCom.Execute SQL
                End If
                cConexCom.Execute "delete from art_serie "
                cConexCom.Execute "delete from art_LOTE "
                DTPicker1 = Date
                DTPicker1.SetFocus
               ' End If
            End If
        End If
    End If
    'Enteros_Positivos KeyAscii, txtNum
End Sub

Function Estado_Valido() As Boolean
    Dim vest As String
    
    vest = Devolver_Dato(1, txtNum, "comovc", "oc_cnumord", False, "oc_csitord")
    Estado_Valido = False
    If vest = "01" Or vest = "03" Then Estado_Valido = True
End Function

Sub Muestra_datos_de_COMOVC()
    Dim strsql, lblRuc As String
    
    Set Adodc1 = New ADODB.Recordset
    
    strsql = "SELECT ccodprove,cdesprove,femision,fentrega,csituacion, " & _
        "ccodmonim FROM imporc WHERE cnumero='" & txtNum & "' and csituacion<>'04'"
    Adodc1.Open strsql, cConexCom, adOpenDynamic, adLockOptimistic
    If Not Adodc1.EOF Then
       lblPro = Adodc1("ccodprove")
       lblProv = Adodc1("cdesprove")
       ' lblRuc =  Devolver_Dato(1, lblPro, "maeprov", "prvccodigo", False, "prvcruc")
       lblEmi = Adodc1("femision")
       lblEnt = Adodc1("fentrega")
    Else
       MsgBox "La Orden de Importación no Tiene Saldos", vbInformation, "Seleccione otra Orden"
       txtNum = ""
    End If
    'lblEst = Adodc1("oc_csitord")
    'lblEsta = Devolver_Dato(1, lblEst, "estado_oc", "est_codigo", False, "est_nombre")
End Sub

Sub Rellena_FlexGrid()
    Dim Adodc2 As ADODB.Recordset
    Dim strsql As String, k As Integer
    
    Set Adodc2 = New ADODB.Recordset
    
'    strsql = "SELECT ccodartic,cdesartic,cunidad,ncantidad,ncantentr " & _
        "FROM impord WHERE cnumero='" & txtNum & "' ORDER BY citem"
     strsql = "SELECT impord.CCODARTIC, impord.CDESARTIC, impord.CUNIDAD, impord.NCANTIDAD, impord.NCANTENTR ,IIf(IIf(AFSERIE='S','S','')='',IIF(AFLOTE='S','L',''),IIf(AFSERIE='S','S','')) AS SER_LOT FROM MAEART INNER JOIN impord ON MAEART.ACODIGO = impord.CCODARTIC WHERE impord.CNUMERO='" & txtNum & "'" & _
              " ORDER BY impord.CITEM "
    Adodc2.Open strsql, cConexCom, adOpenStatic
    
    Do While Not Adodc2.EOF
        k = k + 1
        If k = 1 Then
            Flex1.AddItem Adodc2("ccodartic") & vbTab & Adodc2("cdesartic") & vbTab & _
                Adodc2("cunidad") & vbTab & Format(Adodc2("ncantidad"), "0.00") & _
                vbTab & Format(Adodc2("ncantidad") - Val("0" & Adodc2("ncantentr")), "0.00") & vbTab & "0.00" & vbTab & Adodc2("SER_LOT"), 1
            Flex1.Rows = 2
        Else
            Flex1.AddItem Adodc2("ccodartic") & vbTab & Adodc2("cdesartic") & vbTab & _
                Adodc2("cunidad") & vbTab & Format(Adodc2("ncantidad"), "0.00") & _
                vbTab & Format(Adodc2("ncantidad") - Val("0" & Adodc2("ncantentr")), "0.00") & vbTab & "0.00" & vbTab & Adodc2("SER_LOT")
        End If
        
        Adodc2.MoveNext
    Loop
    Adodc2.Close
End Sub

Private Sub txtTF_Change()
    If lblTF <> "" Then
        lblTF = ""
        txtNTF = ""
        txtNTF.Enabled = False
    End If
End Sub

Private Sub txtTF_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT tdo_tipdoc,tdo_descri FROM tipo_docu"
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Tipo de Documentos"
    'frmReferencia.Inicio
    frmReferencia.Show vbModal
    Adodc2.Close
    
    If vGUtil(1) <> "" Then
        txtTF = vGUtil(1)
        txtTF_KeyPress 13
    End If
End Sub

Private Sub txtTF_GotFocus()
    Enfoque txtTF
End Sub

Private Sub txtTF_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtTF_DblClick
End Sub

Private Sub txtTF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTF = Trim(txtTF)
        If txtTF <> "" Then
            If Not Existe(1, txtTF, "tipo_docu", "tdo_tipdoc", False) Then
                Mensaje = "No existe el Tipo de documento ingresado"
                MsgBox Mensaje, vbExclamation, "Error"
                txtTF.SetFocus
            Else
                lblTF = Devolver_Dato(1, txtTF, "tipo_docu", "tdo_tipdoc", False, _
                    "tdo_descri")
                txtNTF.Enabled = True
                txtNTF.SetFocus
            End If
        Else
            txtAlm.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Private Sub txtTM_Change()
    If lblTM <> "" Then lblTM = ""
End Sub

Private Sub txtTM_DblClick()
    Static Adodc2 As ADODB.Recordset
    Dim strsql As String
    
    Set Adodc2 = New ADODB.Recordset
    
    strsql = "SELECT tt_codmov,tt_descri FROM tabtransa where tt_tipmov='I'"
    Adodc2.Open strsql, cConexCom, adOpenStatic, adLockReadOnly
    
    frmReferencia.Conectar Adodc2, strsql
    frmReferencia.Label1 = "Tipo de Movimiento"
    'frmReferencia.iNICIO
    frmReferencia.Show vbModal
    Adodc2.Close
    
    If vGUtil(1) <> "" Then
        txtTM = vGUtil(1)
        txtTM_KeyPress 13
    End If
End Sub

Private Sub txtTM_GotFocus()
    Enfoque txtTM
End Sub

Private Sub txtTM_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 112 Then txtTM_DblClick
End Sub

Private Sub txtTM_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtTM = Trim(txtTM)
        If txtTM <> "" Then
            If Not Existe(1, txtTM, "tabtransa", "tt_codmov", False) Then
                Mensaje = "El Código de Almacén ingresado no existe"
                MsgBox Mensaje, vbExclamation, "Error"
                txtTM.SetFocus
            Else
                lblTM = Devolver_Dato(1, txtTM, "tabtransa", "tt_codmov", False, "tt_descri")
                Flex1.SetFocus
            End If
        Else
            Flex1.SetFocus
        End If
    Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    End If
End Sub

Function Ingreso_Realizado() As Integer
    Dim i As Integer
    Dim tSal As Double
    Dim tRec As Double
    
    For i = 1 To Flex1.Rows - 1
        tSal = tSal + Val(Flex1.TextMatrix(i, 4))
        tRec = tRec + Val(Flex1.TextMatrix(i, 5))
    Next
    
    If tRec = 0 Then
        Ingreso_Realizado = 0
    ElseIf tRec < tSal Then
        Ingreso_Realizado = 1
    Else
        Ingreso_Realizado = 2
    End If
End Function

Sub VerLotes()
Screen.MousePointer = 11
    If Flex1.TextMatrix(Flex1.Row, 4) = 0 Then
       MsgBox "La Cantidad Ingresada no es correcta", vbCritical, "Aviso.....!"
       Exit Sub
    End If
    
    frmVerlotes.ncant = Flex1.TextMatrix(Flex1.Row, 4)
    frmVerlotes.almacen = txtAlm
    frmVerlotes.cCod = Flex1.TextMatrix(Flex1.Row, 0)
    frmVerlotes.cDesc = Flex1.TextMatrix(Flex1.Row, 1)
    '********************************************************
    frmReglotes.Frame2.Visible = False
    frmReglotes.almacen = txtAlm
    frmReglotes.LoadLotesArti (Flex1.TextMatrix(Flex1.Row, 0))
    frmReglotes.Text1 = Flex1.TextMatrix(Flex1.Row, 0)
    frmReglotes.Label3 = Flex1.TextMatrix(Flex1.Row, 1)
    frmReglotes.Caption = "Seleccione o Adicione el Lote Destino "
    frmReglotes.cmdExitimport.Visible = True
    frmReglotes.cmdsubsalida.Visible = False
    frmReglotes.cmdretorna.Visible = True
    frmReglotes.lblmsg.Visible = True
Screen.MousePointer = 1
    frmReglotes.Show 1
    
End Sub

Function CargaLista(ByVal arCod As String) As Boolean
Dim rs As New ADODB.Recordset
   

rs.Open "SELECT ART_LOTE.ALMA, ART_LOTE.ACODIGO, MAEART.ADESCRI, ART_LOTE.LOTE,art_lote.cantid FROM ART_LOTE INNER JOIN MAEART ON ART_LOTE.ACODIGO = MAEART.ACODIGO where  ART_LOTE.alma='" & txtAlm & "'  and ART_LOTE.acodigo='" & arCod & "'", cConexCom, adOpenStatic, adLockBatchOptimistic
If Not rs.EOF Then
'   frmVerlotes.almacen = RS!alma
    frmVerlotes.Gridlote.Clear
    frmVerlotes.Gridlote.Rows = 0
    frmVerlotes.Gridlote.Cols = 3

    frmVerlotes.cCod = rs!ACODIGO
    frmVerlotes.cDesc = rs!ADESCRI
    frmVerlotes.ncant = Flex1.TextMatrix(Flex1.Row, 4)
    Do While Not rs.EOF
       frmVerlotes.Gridlote.AddItem rs!Lote & Chr(9) & Format(rs!cantid, "###,##0.00") & Chr(9) & ClsTock.SaldoLote(rs!alma, rs!ACODIGO, rs!Lote, cConexCom)
       rs.MoveNext
    Loop
    CargaLista = True
    frmVerlotes.Show 1
Else
    CargaLista = False
    frmVerlotes.Gridlote.Clear
    frmVerlotes.Gridlote.Rows = 0
    frmVerlotes.Gridlote.Cols = 3

End If
rs.Close

End Function

