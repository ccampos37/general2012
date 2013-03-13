VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormEliminaDoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Eliminar de Documentos"
   ClientHeight    =   4830
   ClientLeft      =   1740
   ClientTop       =   1770
   ClientWidth     =   9690
   Icon            =   "FormGuiaRem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   9690
   Begin MSDataGridLib.DataGrid dbg_detalle 
      Height          =   2055
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   3625
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   8
      BeginProperty Column00 
         DataField       =   "CANUMDOC"
         Caption         =   "Num. Doc"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "CAALMA"
         Caption         =   "Almacén"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "CAFECDOC"
         Caption         =   "Fecha Doc."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "CATIPMOV"
         Caption         =   "Tipo Mov."
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column04 
         DataField       =   "CACODMOV"
         Caption         =   "Movimiento"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column05 
         DataField       =   "CACODMON"
         Caption         =   "Moneda"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column06 
         DataField       =   "CACODPRO"
         Caption         =   "Proveedor"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column07 
         DataField       =   "CACODCLI"
         Caption         =   "Cliente"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         ScrollBars      =   3
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
         BeginProperty Column02 
         EndProperty
         BeginProperty Column03 
         EndProperty
         BeginProperty Column04 
         EndProperty
         BeginProperty Column05 
         EndProperty
         BeginProperty Column06 
         EndProperty
         BeginProperty Column07 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command8 
      Caption         =   "&Salir"
      Height          =   675
      Left            =   5160
      Picture         =   "FormGuiaRem.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3960
      Width           =   775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      Height          =   675
      Left            =   3840
      Picture         =   "FormGuiaRem.frx":0D0C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3960
      Width           =   775
   End
   Begin VB.Frame Frame1 
      Caption         =   "Buscar"
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   2664
      TabIndex        =   7
      Top             =   144
      Width           =   6315
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   480
         Width           =   3255
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "FormGuiaRem.frx":114E
         Left            =   1440
         List            =   "FormGuiaRem.frx":1158
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   3255
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   285
         Left            =   1440
         TabIndex        =   15
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   63111169
         CurrentDate     =   37928
      End
      Begin VB.Label Label1 
         Caption         =   "Filtro"
         Height          =   255
         Left            =   480
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Indice"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipo"
      ForeColor       =   &H80000007&
      Height          =   1455
      Left            =   90
      TabIndex        =   10
      Top             =   144
      Width           =   2175
      Begin VB.OptionButton Option4 
         Caption         =   "Guia de Compra"
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1185
         Width           =   1770
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nota de Ingreso"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nota de Salida"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   570
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Guias"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   855
         Width           =   1335
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   90
      TabIndex        =   11
      Top             =   144
      Visible         =   0   'False
      Width           =   2175
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "FormGuiaRem.frx":116B
         Top             =   480
         Width           =   480
      End
      Begin VB.Label Label3 
         Caption         =   "GUIAS   DE REMISION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   648
         TabIndex        =   12
         Top             =   480
         Width           =   1212
      End
   End
End
Attribute VB_Name = "FormEliminaDoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim G_SQL           As String
Dim Tipo            As String
Dim Estado          As String
Dim cadena          As String
Dim Fecha           As Date
Dim serie_lote      As String
Dim noprocede       As Boolean
Dim estimp          As String
Dim cCodprovee      As String
Dim rSTD01          As New ADODB.Recordset
Dim rSTD02          As New ADODB.Recordset
Dim Adodc1          As ADODB.Recordset
Private Sub Combo1_Click()

    If Me.Combo1.ListIndex = 0 Then
        Me.Text1.Visible = True
        Me.DTPicker1.Visible = False
    Else
        Me.Text1.Visible = False
        Me.DTPicker1.Visible = True
    End If
  If Combo1.text = "Numero" Then
      If VGLadrillera Then
         G_SQL = "SELECT *, " & _
         "(CASE casitgui WHEN 'F' THEN 'FACTURADO' ELSE " & _
         "(CASE CASITGUI WHEN 'A' THEN 'ANULADO' ELSE " & _
         "(CASE CASITGUI WHEN 'P' THEN 'PENDIENTE' ELSE 'VIGENTE' END) END) END) AS COL1, CAESTIMP " & _
         " FROM MovAlmCab where CAALMA='" & VGAlma & "' AND  (CATD='" & Tipo & "' OR CATD='GF') " & _
         " ORDER BY CANUMDOC "
      Else
         G_SQL = "SELECT *, " & _
         " (CASE casitgui WHEN 'F' THEN 'FACTURADO' ELSE " & _
         " (CASE CASITGUI WHEN 'A' THEN 'ANULADO' ELSE " & _
         " (CASE CASITGUI WHEN 'P' THEN 'PENDIENTE' ELSE " & _
         " (CASE CASITGUI WHEN 'V' THEN 'VIGENTE' ELSE 'NULL' END) END) END) END) AS COL1,CAESTIMP " & _
         " FROM MovAlmCab WHERE CAALMA='" & VGAlma & "' AND  (CATD='" & Tipo & "' OR CATD='GF') " & _
         " and (CASITGUI<>'E'and CASITGUI<>'F') " & IIf(Option4.Value, " and CARFTDOC='GC'", " " & _
         " and (CARFTDOC<>'GC' OR CARFTDOC=NULL) ") & " ORDER BY CANUMDOC "
         'ORIGINAL
         '" and (CASITGUI<>'E' and CASITGUI<>'F') " & IIf(Option4.Value, " and CARFTDOC='GC'", "
      End If
  Else
      If VGLadrillera Then
         G_SQL = "SELECT * ," & _
         " (CASE casitgui WHEN 'F' THEN 'FACTURADO' ELSE " & _
         " (CASE CASITGUI WHEN 'A' THEN 'ANULADO' ELSE " & _
         " (CASE CASITGUI WHEN 'P' THEN 'PENDIENTE' ELSE 'VIGENTE' END) END) END)AS COL1, CAESTIMP " & _
         " FROM MovAlmCab where CAALMA='" & VGAlma & "' AND  (CATD='" & Tipo & "' OR CATD='GF') " & _
         " ORDER BY CAFECDOC"
      Else
         G_SQL = "SELECT * ,  (CASE casitgui WHEN 'F' THEN 'FACTURADO' ELSE " & _
         " (CASE CASITGUI WHEN 'A' THEN 'ANULADO' ELSE " & _
         " (CASE CASITGUI WHEN 'P' THEN 'PENDIENTE' ELSE 'VIGENTE' END) END) END) AS COL1,CAESTIMP " & _
         " FROM MovAlmCab where CAALMA='" & VGAlma & "' AND  (CATD='" & Tipo & "' OR CATD='GF') " & _
         " and (CASITGUI<>'E' and CASITGUI<>'F') " & IIf(Option4.Value, " and CARFTDOC='GC'", " " & _
         " and (CARFTDOC<>'GC' OR CARFTDOC=NULL) ") & "  ORDER BY CAFECDOC"
         'ORIGINAL
         '" and (CASITGUI<>'E'  and CASITGUI<>'F') " & IIf(Option4.Value, " and CARFTDOC='GC'",
      End If
  End If
  If Adodc1.State = adStateOpen Then Adodc1.Close
  Adodc1.Open G_SQL, cConexCom, adOpenDynamic, adLockOptimistic
  Set Me.dbg_detalle.DataSource = Adodc1
  Me.dbg_detalle.Refresh
  Label1.Caption = Combo1.text
  End Sub

Private Sub Command1_Click()
  Unload Me
End Sub
'Aceptar
Private Sub Command2_Click()
   Dim Num      As String
   Dim rpta     As Integer
   Dim transa   As String
   Dim RSQL     As String
  
   On Error GoTo Err
   estimp = ""
   cCodprovee = ""
   If Adodc1.RecordCount = 0 Then Exit Sub
   Fecha = Adodc1("CAFECDOC")
   Num = Adodc1("CANUMDOC")
   transa = Adodc1("CACODMOV")
   estimp = IIf(IsNull(Adodc1("CASITGUI")), "V", Adodc1("CASITGUI"))
   cCodprovee = Adodc1("CACODPRO") & " "
   'Si el documento
   If Adodc1("CAESTIMP") = "I" Then
         If VGElimina Then
            Rem MVV MsgBox "La Guia no se puede Eliminar, porque ya ha sido Impresa", vbInformation, "Información"
            rpta = MsgBox("Este documento este ha sido Impreso, ¿desea eliminar de todas maneras?", vbYesNo + vbInformation, "Información")
            If rpta <> vbYes Then
                Exit Sub
            End If
        End If
    End If
   
    If estimp = "F" And Tipo = "GS" And VGElimina Then
        MsgBox "El documento: " & Num & " no se puede Eliminar  ", vbCritical + vbOKOnly, "Aviso"
        Exit Sub
    End If
   
   If transa = "TD" And Adodc1("CAtd") <> "NI" And VGElimina Then 'Or transa = "GF" cambio Luchito
       
       rpta = MsgBox("Seguro de Eliminar el documento ? " & Num, vbQuestion + vbOKCancel, "Aviso")
       If rpta = vbOK Then
          'RMM********************************
           Call EliminaDocumento(Adodc1!CAALMA, Adodc1("CAtd"), Adodc1!CANUMDOC)
          '**********************************
          Combo1_Click
       End If
      
      Exit Sub
   Else
       If transa = "TD" And VGElimina Then 'GUIAS POR TRANSFERENCIA ANULADAS
          MsgBox "Tipo de documento: " & transa & " Solo se puede Eliminar desde el Almacén Origen ", vbExclamation, "Aviso"
          Exit Sub
       End If
       
       If transa = "TD" And Tipo = "GS" And Not VGElimina Then
          If estimp <> "A" Then
             If MsgBox("Seguro de Anular el documento ? " & Num, vbQuestion + vbOKCancel, "Aviso") = vbOK Then
                Call AnulaDocumento(Adodc1!CAALMA, Adodc1("CAtd"), Adodc1!CANUMDOC)
             End If
             dbg_detalle.Refresh
             Combo1_Click
             dbg_detalle.Refresh
          Else
             MsgBox "El Documento ya esta Anulado ..... ", vbExclamation, "Aviso"
             Exit Sub
          End If
          Exit Sub
       End If

   End If
   'verifica_doc_eli         '
   If Adodc1("CASITGUI") = "A" And VGElimina = False Then
          MsgBox "Tipo de documento  ya ha sido Anulado  ", vbExclamation, "Aviso"
          Exit Sub
   End If
   If Not sihaystk(Tipo, Num) Then Exit Sub
   
   
   
   'CONSIDICION QUE CONPRUEBA SI LA GUIA DE COMPRA ESTA FACTURADA SI O NO
'   If Option4.Value Or funcDevuelveValor("SELECT count(*) FROM COMGUICAB WHERE CCNUMSER+CCNUMDOC='" & Adodc1("CARFNDOC") & "' AND CCCODPRO='" & Adodc1("CACODPRO") & "' AND CCALMA ='" & VGAlma & "'", cConexCom) > 0 Then
'      If funcDevuelveValor("SELECT CCESTADO FROM COMGUICAB WHERE CCNUMSER+CCNUMDOC='" & Adodc1("CARFNDOC") & "' AND CCCODPRO='" & Adodc1("CACODPRO") & "' AND CCALMA ='" & VGAlma & "'", cConexCom) = "F" Then
'        MsgBox "No se puede Eliminar este registro,la guia esta Facturada ", vbCritical
'        Exit Sub
'      End If
'   End If


'verificar si esta cerrado elñ m,es
   If Adodc1("cacierre") Then
      MsgBox "El Mes esta Cerrado para este Documento  ", vbCritical
      Exit Sub
   End If
   If VGElimina Then
        rpta = MsgBox("Seguro de Eliminar el documento ? " & Num, vbQuestion + vbOKCancel, "Aviso")
     Else
        rpta = MsgBox("Seguro de Anular el documento ? " & Num, vbQuestion + vbOKCancel, "Aviso")
   End If
   If rpta = vbOK Then
        '
        'If noprocede Then Exit Sub
        If Not VGElimina Then
                 RSQL = "update  movalmcab  set casitgui = '" & Estado & "' where catd='GS' and caalma= '" & VGAlma & "' and canumdoc ='" & Num & "'  "
                 estimp = "A"
                 cConexCom.Execute RSQL
                 Call Descarga(Tipo, Num)
                 MsgBox "Se Anuló el documento", vbInformation, "Aviso"
        End If
        
        If estimp <> "A" Then
           Call Descarga(Tipo, Num)
        Else
            If VGElimina Then
               cConexCom.Execute "delete from movalmdet where  DEALMA ='" & VGAlma & "' AND DETD = '" & Tipo & "' AND DENUMDOC ='" & Num & "'"
               cConexCom.Execute "delete from movalmcab where  CAALMA ='" & VGAlma & "' AND CATD = '" & Tipo & "' AND CANUMDOC ='" & Num & "' "
            End If
        End If
        
        If Option1.Value = True Then  'Tipo = "NI"
                Call Option1_Click
        ElseIf Option2.Value = True Then 'Tipo = "NS"
                Call Option2_Click
        Else
                Call Option3_Click
        End If
                                                      
        Combo1_Click
        
    End If
    Exit Sub
Err:
    MsgBox Err.Description, vbInformation, "Aviso"
    Exit Sub
    Resume
End Sub

Private Sub Command8_Click()
  Unload Me
End Sub

Private Sub dbg_detalle_Click()
 If dbg_detalle.Row <> 0 Then
    Command2.Enabled = True
 End If
End Sub

Private Sub Form_Load()
    Me.Height = 5205
    Me.Width = 10050
    
    Set Adodc1 = New ADODB.Recordset
    
    Combo1.ListIndex = 0
    Label1.Caption = Combo1.text
  
    
    Init_ControlDataGrid dbg_detalle
    central Me
    If VGElimina Then
        FormEliminaDoc.Caption = "Eliminación de Documentos"
        Option1.Value = True
    Else
        Option3.Value = True
        FormEliminaDoc.Caption = "Anulación de Guía"
        Estado = "A"
        Frame2.Visible = False
        Frame3.Visible = True
        Call Option3_Click
    End If
End Sub

Private Sub Option1_Click()
  If FormEliminaDoc.Caption = "Anulación de Guía" Then Exit Sub
  'Elimina Documentos
   G_SQL = "SELECT *, " & _
   " (CASE casitgui WHEN 'F' THEN 'FACTURADO' ELSE " & _
   " (CASE CASITGUI WHEN 'A' THEN 'ANULADO' ELSE " & _
   " (CASE CASITGUI WHEN 'V' THEN 'VIGENTE' ELSE '' END) END) END) AS COL1,CAESTIMP,CACODMOV " & _
   " FROM MovAlmCab where CAALMA='" & VGAlma & "' and  CATD='NI' and CASITGUI<>'E'  AND " & _
   "  NOT(CARFTDOC='GC')  ORDER BY CANUMDOC "
   If Adodc1.State = adStateOpen Then Adodc1.Close
Screen.MousePointer = 11
  Adodc1.Open G_SQL, cConexCom, adOpenKeyset, adLockOptimistic 'casitua vigente que nosea anulado
  Set Me.dbg_detalle.DataSource = Adodc1
  Me.dbg_detalle.Refresh
Screen.MousePointer = 1
  Tipo = "NI"
  Call Combo1_Click
End Sub
Private Sub Option2_Click()
    G_SQL = "SELECT *," & _
    " (CASE casitgui WHEN 'F' THEN 'FACTURADO' ELSE " & _
    " (CASE CASITGUI WHEN 'A' THEN 'ANULADO' ELSE " & _
    " (CASE CASITGUI WHEN 'V' THEN 'VIGENTE' ELSE '' END) END) END) AS COL1,CAESTIMP,CACODMOV " & _
    " FROM MovAlmCab where CAALMA='" & VGAlma & "' and   CATD='NS' and CASITGUI<>'E'   " & _
    " ORDER BY CANUMDOC "
    If Adodc1.State = adStateOpen Then Adodc1.Close
Screen.MousePointer = 11
    Adodc1.Open G_SQL, cConexCom, adOpenDynamic, adLockOptimistic
    Set Me.dbg_detalle.DataSource = Adodc1
    Me.dbg_detalle.Refresh
    Tipo = "NS"
    Call Combo1_Click
Screen.MousePointer = 1
End Sub
Private Sub Option3_Click()
    G_SQL = "SELECT *," & _
    " (CASE casitgui WHEN 'F' THEN 'FACTURADO' ELSE " & _
    " (CASE CASITGUI WHEN 'A' THEN 'ANULADO' ELSE " & _
    " (CASE CASITGUI WHEN 'P' THEN 'PENDIENTE' ELSE 'VIGENTE' END) END) END) AS COL1,CAESTIMP,CACODMOV " & _
    " FROM MovAlmCab where CAALMA='" & VGAlma & "' and   (CATD='GS' OR CATD='GV')  and CASITGUI<>'E' " & _
    " ORDER BY CANUMDOC "
    If Adodc1.State = adStateOpen Then Adodc1.Close
    Screen.MousePointer = 11
    Adodc1.Open G_SQL, cConexCom, adOpenDynamic, adLockOptimistic
    Set Me.dbg_detalle.DataSource = Adodc1
    Me.dbg_detalle.Refresh
    Tipo = "GS"
    Call Combo1_Click
    Screen.MousePointer = 1
End Sub

Private Sub Option4_Click()
  G_SQL = "SELECT *, (CASE casitgui WHEN 'F' THEN 'FACTURADO' ELSE (CASE CASITGUI WHEN 'A' THEN 'ANULADO' ELSE (CASE CASITGUI WHEN 'V' THEN 'VIGENTE' ELSE '' END) END) END) AS COL1,CAESTIMP,CACODMOV " & _
  " FROM MovAlmCab where CAALMA='" & VGAlma & "' and  CATD='NI' and CASITGUI<>'E'  " & _
  " AND (CARFTDOC='GC')  ORDER BY CANUMDOC " 'casitua vigente que nosea anulado
    If Adodc1.State = adStateOpen Then Adodc1.Close
Screen.MousePointer = 11
  Adodc1.Open G_SQL, cConexCom, adOpenDynamic, adLockOptimistic
  Set Me.dbg_detalle.DataSource = Adodc1
  dbg_detalle.Refresh
  Tipo = "NI"
  Call Combo1_Click
  Screen.MousePointer = 1
End Sub
Private Sub Text1_Change()
On Error GoTo Mensaje
    Dim ncar As String
    Dim criterio As String
    
    ncar = Str$(Len(Text1.text))
    If Combo1.text = "Numero" Then
        Rem MVV criterio = "MID$(CANUMDOC,1," + ncar + ") = " & Chr$(34) + Text1.text + Chr$(34)
        criterio = " CANUMDOC LIKE '" & Text1.text & "*'"
    Else
        Rem mvv criterio = "MID$(CAFECDOC,1," + ncar + ") =  " & Chr$(34) + Text1.text + Chr$(34)     '     #" & Text1.text & "#"
        criterio = "CAFECDOC =  " & DateSQL(Me.DTPicker1.Value) & ""
    End If
    Adodc1.Filter = criterio
Mensaje:
    Captura_error
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'   Dim ncar As String
'   Dim criterio  As String
'   ncar = Str$(Len(Text1.text))
'   If Combo1.text = "Numero" Then
'        criterio = "MID$(CANUMDOC,1," + ncar + ") = " & Chr$(34) + Text1.text + Chr$(34)
'   Else
'        criterio = "MID$(CAFECDOC,1," + ncar + ") = " & Chr$(34) + Text1.text + Chr$(34)
'   End If
'   Data1.Recordset.FindFirst criterio
End Sub

Function sihaystk(doc As String, NumDoc As String)
  Dim AdoReg1       As ADODB.Recordset
  Dim RSQL          As String
  Dim suma          As Boolean
  Dim verdad        As Boolean
  Dim sSqlCad       As String
  
  
  RSQL = "select  DECODIGO, DECANTID from MovAlmDet   where DEALMA = '" & VGAlma & "'  and DETD= '" & doc & "' AND   DENUMDOC= '" & NumDoc & "' and decodigo<>'TEXTO'"
  Set AdoReg1 = New ADODB.Recordset
  AdoReg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
  If AdoReg1.RecordCount <> 0 Then
      
  End If
  If doc = "NI" Then
        suma = False
  Else
        suma = True
  End If
  sihaystk = True
  While Not AdoReg1.EOF And sihaystk
        If Not verificastk(AdoReg1(0), AdoReg1(1), suma) Then
                MsgBox "El articulo " & AdoReg1(0) & " no tiene suficiente Stock para la Eliminación, " & Chr(13) & "No se puede eliminar el documento ", vbExclamation, "Eliminar"
                sihaystk = False
        End If
        AdoReg1.MoveNext
  Wend
  AdoReg1.Close
  
If doc <> "NI" Then
   sSqlCad = "SELECT  ST.STSALMA,ST.STSCODIGO,ST.STSSERIE,ST.STSSKDIS FROM MOVALMDET "
   sSqlCad = sSqlCad & " MD INNER JOIN STKSERI ST ON MD.DECODIGO=ST.STSCODIGO "
   sSqlCad = sSqlCad & "AND MD.DESERIE=ST.STSSERIE Where MD.DEALMA='" & VGAlma & "' AND MD.DETD "
   sSqlCad = sSqlCad & "='" & doc & "' AND MD.DENUMDOC='" & NumDoc & "' AND ST.STSSKDIS = 1"
   Set AdoReg1 = New ADODB.Recordset
   AdoReg1.Open sSqlCad, cConexCom, adOpenStatic, adLockOptimistic
   If AdoReg1.RecordCount > 0 Then
       MsgBox "Existen articulos seriados que no pueden ser eliminados", vbCritical
       sihaystk = False
   End If
End If

End Function
Public Sub Descarga(doc As String, NumDoc As String)
On Error GoTo Mensaje
  Dim AdoReg1 As ADODB.Recordset
  Dim RSQL, csql, SERLOTE As String
  Dim suma As Boolean
  Dim dato As String
  Dim codigo() As String
  Dim n, X As Long
  Dim OrdImport As String
  Dim NumFactura As String
  Dim NroPed As String
  Dim NroGuiaProv As String
  OrdImport = ClsTDoc.EsImportacion(VGAlma, doc, NumDoc, cConexCom)
  If VGLadrillera Then NumFactura = ClsTDoc.EsDespachoFactura(VGAlma, doc, NumDoc, cConexCom)
  NroPed = ClsTDoc.TienePedido(VGAlma, "GS", NumDoc, cConexCom)
  NroGuiaProv = ClsTDoc.NroGuiaproveedor(VGAlma, NumDoc, cConexCom)
  noprocede = False
  If NroGuiaProv <> "" Then
     RSQL = "SELECT * from COMGUICAB Where CCNUMSER+CCNUMDOC='" & NroGuiaProv & "' AND CCCODPRO = '" & cCodprovee & "' AND  CCESTADO<>'F'"
     Set AdoReg1 = New ADODB.Recordset
     AdoReg1.Open RSQL, cConexCom, adOpenStatic, adLockOptimistic
     If AdoReg1.EOF Then
        MsgBox "La Guia de Compra Correspondiente a esta Nota de Ingreso ya esta Facturada" & Chr(10) & "Usted No puede Eliminar Esta Nota de Ingreso...!", vbExclamation, "Aviso.."
        Exit Sub
     End If
     AdoReg1.Close
  End If
  
  '**********************
  Erase codigo
  '**********************
   
  RSQL = "select  DECODIGO, DECANTID,DESERIE,DELOTE,DEITEMI from MovAlmDet   where DEALMA = '" & VGAlma & "'  and DETD= '" & doc & "' AND   DENUMDOC= '" & NumDoc & "'"
  Set AdoReg1 = New ADODB.Recordset
  AdoReg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
  If AdoReg1.RecordCount = 0 Then
       MsgBox "No hay detalle del documento", vbCritical, "Aviso"
       noprocede = True
       GoTo fin
  Else
     ReDim codigo(AdoReg1.RecordCount) As String
  End If
      
X = 0
'*RMM***********************
While Not AdoReg1.EOF
      X = X + 1
      '*RMM***********************
      If OrdImport <> "" Then ClsTDoc.RestauraDocImport OrdImport, AdoReg1!DECODIGO, AdoReg1!DECANTID, AdoReg1!DEITEMI, cConexCom
      If NumFactura <> "" Then ClsTDoc.RestauraDespachoFactura NumFactura, AdoReg1!DECODIGO, AdoReg1!DECANTID, cConexCom
      If NroPed <> "" Then ClsTDoc.RestaurarPedido NroPed, AdoReg1!DECODIGO, AdoReg1!DECANTID, cConexCom
      
      codigo(X) = AdoReg1!DECODIGO
     AdoReg1.MoveNext
Wend

If OrdImport <> "" Then
   ClsTDoc.DefineEstado OrdImport, cConexCom
End If
 
If NroGuiaProv <> "" Then
   ClsTDoc.EliminoGuiaProveeCompra NroGuiaProv, cCodprovee, cConexCom
End If
 
  AdoReg1.Close
  
  If VGElimina Then
            csql = "delete from movalmdet where  DEALMA ='" & VGAlma & "' AND DETD = '" & doc & "' AND DENUMDOC ='" & NumDoc & "'"
            cConexCom.Execute csql
  End If
fin:
  If VGElimina Then
           csql = "delete from movalmcab where  CAALMA ='" & VGAlma & "' AND CATD = '" & doc & "' AND CANUMDOC ='" & NumDoc & "' "
           cConexCom.Execute csql
  End If

SERLOTE = ""

For n = 1 To X
    SERLOTE = ClsTock.EsSerie_Lote(codigo(n), cConexCom)
    If SERLOTE = "S" Then ClsTock.CalculaStockSerie VGAlma, codigo(n)
    If SERLOTE = "L" Then ClsTock.CalculaStockLOTE VGAlma, codigo(n)
    ClsTock.CalculaSaldoNoValorizado VGAlma, codigo(n), Fecha  'Actualiza stkart , Moresmes
Next
   
'*RMM***********************
  ClsTDoc.CorrigueNumeracion VGAlma, cConexCom
  Combo1_Click
  Me.dbg_detalle.Refresh
Exit Sub
Mensaje:
    Captura_error
End Sub

Public Sub ObtenerCantDoc(doc As String, NumDoc As String)

End Sub

Function verificastk(Cod As String, CANTIDAD As Double, suma As Boolean)
  Dim rs As ADODB.Recordset
  Dim RSQL As String
  Dim dato As Double
  verificastk = False
  
  RSQL = "select n.STSKDIS from  StkArt n WHERE n.STALMA = '" & VGAlma & "'  and  n.STCODIGO= '" & Cod & "' "
  Set rs = New ADODB.Recordset
  rs.Open RSQL, cConexCom, adOpenStatic, adLockOptimistic
  
  If rs.EOF Then
       MsgBox "No hay el código articulo en almacen", vbCritical
       Exit Function
  End If
  verificastk = True
  If suma Then    'DESCARGAR AL MOREMES
     dato = rs(0) + CANTIDAD  'revisar
  Else
     dato = rs(0) - CANTIDAD
  End If
  If dato < 0 Then verificastk = False
End Function
Public Sub buscarstk(Cod As String, CANTIDAD As Double, suma As Boolean, tieneserie As String)
  Dim rs As ADODB.Recordset
  Dim RSQL As String
  Dim dato As Double
  
  RSQL = "select n.STSKDIS from  StkArt n WHERE n.STALMA = '" & VGAlma & "'  and  n.STCODIGO= '" & Cod & "' "
  Set rs = New ADODB.Recordset
  rs.Open RSQL, cConexCom, adOpenStatic, adLockOptimistic
  
  If rs.EOF Then
       MsgBox "No hay dicho articulo en almacen", vbCritical
       Exit Sub
  End If
  If suma Then
        dato = rs(0) + CANTIDAD  'revisar
  Else
        dato = rs(0) - CANTIDAD
  End If
  If tieneserie = "S" Then actserie (Cod)
  If tieneserie = "N" Then actlote Cod, CANTIDAD
  If dato <> 0 Then
       RSQL = "Update STKART set STSKDIS = " & dato & " where STALMA = '" & VGAlma & "'  and  STCODIGO= '" & Cod & "' "
  Else
       RSQL = "Update STKART set STSKDIS = " & dato & "  , STKPREULT= " & dato & "  , STKPREPRO=  " & dato & "  where STALMA = '" & VGAlma & "'  and  STCODIGO= '" & Cod & "' "
  End If
  cConexCom.Execute RSQL
  rs.Close
End Sub
Private Sub actlote(codigo As String, CANTIDAD As Double)
Dim uSql As String
Dim nuevo_stk As Double
Dim RSQL As String
Dim rs As ADODB.Recordset
    
    RSQL = "select STSLKDIS FROM STKLOTE where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSLOTE= '" & serie_lote & "'" '
    Set rs = New ADODB.Recordset
    rs.Open RSQL, cConexCom, adOpenStatic, adLockOptimistic
    
    If Not rs.EOF Then
       If Tipo = "NI" Then
         nuevo_stk = rs(0) - CANTIDAD
       Else
         nuevo_stk = rs(0) + CANTIDAD
       End If
       uSql = "Update STKLOTE set STSLKDIS = " & nuevo_stk & " WHERE  STSALMA= '" & VGAlma & "' and STSCODIGO='" & codigo & "'AND STSLOTE='" & serie_lote & "'"
       cConexCom.Execute uSql
    End If
End Sub

Private Sub actserie(codigo As String)
Dim uSql As String
Dim Serie As String
Dim VALOR As Integer
Dim rs As ADODB.Recordset
Dim RSQL As String
    
RSQL = "select STSSKDIS FROM STKSERI where   STSALMA= '" & VGAlma & "' and  STSCODIGO= '" & codigo & "' and STSSERIE= '" & serie_lote & "'" '
Set rs = New ADODB.Recordset
rs.Open RSQL, cConexCom, adOpenStatic, adLockOptimistic

If Not rs.EOF Then
        VALOR = IIf(Tipo = "NI", 0, 1)
        uSql = "update STKSERI set STSSKDIS = " & VALOR & " WHERE  STSALMA='" & VGAlma & "' and STSCODIGO='" & codigo & "'AND STSSERIE='" & serie_lote & "'"
        cConexCom.Execute uSql
End If
End Sub

Private Sub actvalmes(CANTIDAD As Double, Tipo As String)
  Dim criterio As String
  Dim AdoReg1 As ADODB.Recordset
  Dim RSQL As String
  Dim Cantent As Double
  Dim Cantsal As Double
  Dim mespro As String
  Dim uSql As String
  
  mespro = Year(Fecha) & Format(Month(Fecha), "00")
  RSQL = "select SMCANENT,SMCANSAL FROM MoResMes where  SMALMA = '" & VGAlma & "' and SMMESPRO = '" & mespro & "' AND  SMCODIGO= '" & cadena & "'"  '
   
  Set AdoReg1 = New ADODB.Recordset
  AdoReg1.Open RSQL, cConexCom, adOpenDynamic, adLockOptimistic
  If AdoReg1.RecordCount <> 0 Then
      If Tipo = "NI" Then
            Cantent = AdoReg1(0) - CANTIDAD
            uSql = "Update MoResMes set SMCANENT = " & Cantent & " where SMALMA='" & VGAlma & "' and   SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      Else
            Cantent = AdoReg1(0)
            Cantsal = AdoReg1(1) - CANTIDAD
            uSql = "Update MoResMes set SMCANSAL = " & Cantsal & ",SMCANENT = " & Cantent & "  where SMALMA='" & VGAlma & "' and  SMCODIGO ='" & cadena & "' AND SMMESPRO ='" & mespro & "' "
      End If
  Else
      If Tipo = "NI" Then
            Cantent = CANTIDAD
            Cantsal = 0
      Else
            Cantent = CANTIDAD
            Cantsal = 0
      End If
      uSql = "insert into MoResMes (SMALMA,SMCODIGO,SMMESPRO,SMCANENT,SMCANSAL,SMSALDOINI) VALUES ('" & VGAlma & "','" & cadena & "','" & mespro & "' ," & Cantent & "," & Cantsal & ",0) "
  End If
  cConexCom.Execute uSql
  AdoReg1.Close
End Sub
'ROBERTO MAZA MILLA 14/07/2001
Sub EliminaDocumento(ByVal AlmaActu As String, ByVal Tdoc As String, ByVal NumDoc As String)
On Error GoTo Mensaje
Dim SQL, AlmaRefe, DocRefe, NumRefe As String
Dim SERLOTE As String
Dim alma, TipDoc, NDoc As String
Set rSTD01 = New ADODB.Recordset
Set rSTD02 = New ADODB.Recordset
Dim codigo() As String
Dim nReg, nCont, n As Long


'ESTE RECORDSET CARGA LA TRANSACCION ORIGEN
SQL = " SELECT CAALMA,CARFALMA,CANUMDOC,CATD,CARFTDOC,CACODMOV,CAHORA,CAFECDOC,DECODIGO " & _
     " FROM MovAlmDet AS A INNER JOIN MovAlmCab AS B ON (B.CAALMA = A.DEALMA) AND (B.CATD = A.DETD) AND (B.CANUMDOC = A.DENUMDOC) " & _
     " WHERE CAALMA = '" & AlmaActu & "'And CATD='" & Tdoc & "' AND CACODMOV='TD' AND CANUMDOC='" & NumDoc & "' AND CACIERRE=0 AND (CASITGUI='A' OR CASITGUI='V')"
rSTD01.Open SQL, cConexCom, adOpenStatic

nReg = rSTD01.RecordCount: nCont = 0
ReDim codigo(nReg) As String
Do While Not rSTD01.EOF
   nCont = nCont + 1
   codigo(nCont) = rSTD01!DECODIGO
   rSTD01.MoveNext
Loop
rSTD01.Close

'ESTA CONSULTA ES EN CASO DE UN DOCUMENTO SIN DETALLE
SQL = " SELECT CAALMA,CARFALMA,CANUMDOC,CATD,CARFTDOC,CACODMOV,CAHORA,CAFECDOC " & _
     " FROM MovAlmCab WHERE CAALMA = '" & AlmaActu & "'And CATD='" & Tdoc & "' AND CACODMOV='TD' AND CANUMDOC='" & NumDoc & "' AND CACIERRE=0 AND (CASITGUI='A' OR CASITGUI='V')"
rSTD01.Open SQL, cConexCom, adOpenStatic
'****Restaura Stock Actual
If Not rSTD01.EOF Then
'****Para ubicar Documento en El Almacen Destino
   AlmaRefe = rSTD01!CARFALMA
   NumRefe = rSTD01!CANUMDOC
End If
rSTD01.Close

SQL = " SELECT CAALMA,CARFALMA,CANUMDOC,CATD,CARFTDOC,CACODMOV,CAHORA,CAFECDOC " & _
     " FROM MovAlmCab WHERE CAALMA = '" & AlmaRefe & "'And CATD='NI' AND CACODMOV='TD' AND CARFNDOC ='" & NumRefe & "' AND CACIERRE=0 AND (CASITGUI='A' OR CASITGUI='V')"
rSTD01.Open SQL, cConexCom, adOpenStatic
If Not rSTD01.EOF Then
'****Para ubicar Documento en El Almacen Destino
   NumRefe = rSTD01!CANUMDOC
Else
   NumRefe = ""
End If
rSTD01.Close


'ELIMINO ORIGEN
ClsTDoc.EliminoCabezera AlmaActu, Tdoc, NumDoc, cConexCom
ClsTDoc.EliminoDetalle AlmaActu, Tdoc, NumDoc, cConexCom
'ELIMINO DESTINO
If AlmaRefe <> "" Then
   ClsTDoc.EliminoCabezera AlmaRefe, "NI", NumRefe, cConexCom
   ClsTDoc.EliminoDetalle AlmaRefe, "NI", NumRefe, cConexCom
End If

For n = 1 To nCont
    SERLOTE = ClsTock.EsSerie_Lote(codigo(n), cConexCom)
    If SERLOTE = "S" Then
       ClsTock.CalculaStockSerie AlmaActu, codigo(n)
       If AlmaRefe <> "" Then ClsTock.CalculaStockSerie AlmaRefe, codigo(n)
    End If
    
    If SERLOTE = "L" Then
       ClsTock.CalculaStockLOTE AlmaActu, codigo(n)
       If AlmaRefe <> "" Then ClsTock.CalculaStockLOTE AlmaRefe, codigo(n)
   End If
   
    ClsTock.CalculaSaldoNoValorizado AlmaActu, codigo(n), Now   'Actualiza stkart , Moresmes
    If AlmaRefe <> "" Then ClsTock.CalculaSaldoNoValorizado AlmaRefe, codigo(n), Now
Next

    ClsTDoc.CorrigueNumeracion AlmaActu, cConexCom
    If AlmaRefe <> "" Then ClsTDoc.CorrigueNumeracion AlmaRefe, cConexCom

Exit Sub
Mensaje:
     Captura_error
End Sub
'ROBERTO MAZA MILLA 14/07/2001
Sub AnulaDocumento(ByVal AlmaActu As String, ByVal Tdoc As String, ByVal NumDoc As String)
Dim SQL, AlmaRefe, DocRefe, NumRefe As String
Dim SERLOTE As String
Dim alma, TipDoc, NDoc As String
Set rSTD01 = New ADODB.Recordset
Set rSTD02 = New ADODB.Recordset
Dim codigo() As String
Dim nReg, nCont, n As Long
On Local Error GoTo ERRAR
'ESTE RECORDSET CARGA LA TRANSACCION ORIGEN
SQL = " SELECT CAALMA,CARFALMA,CANUMDOC,CATD,CARFTDOC,CACODMOV,CAHORA,CAFECDOC,DECODIGO " & _
     " FROM MovAlmDet AS A INNER JOIN MovAlmCab AS B ON (B.CAALMA = A.DEALMA) AND (B.CATD = A.DETD) AND (B.CANUMDOC = A.DENUMDOC) " & _
     " WHERE CAALMA = '" & AlmaActu & "'And CATD='" & Tdoc & "' AND CACODMOV='TD' AND CANUMDOC='" & NumDoc & "' AND CACIERRE=0 AND (CASITGUI='A' OR CASITGUI='V')"
rSTD01.Open SQL, cConexCom, adOpenStatic
nReg = rSTD01.RecordCount: nCont = 0
ReDim codigo(nReg) As String

Do While Not rSTD01.EOF
   nCont = nCont + 1
   codigo(nCont) = rSTD01!DECODIGO
   rSTD01.MoveNext
Loop
rSTD01.Close
'ESTA CONSULTA ES EN CASO DE UN DOCUMENTO SIN DETALLE
SQL = " SELECT CAALMA,CARFALMA,CANUMDOC,CATD,CARFTDOC,CACODMOV,CAHORA,CAFECDOC " & _
     " FROM MovAlmCab WHERE CAALMA = '" & AlmaActu & "'And CATD='" & Tdoc & "' AND CACODMOV='TD' AND CANUMDOC='" & NumDoc & "' AND CACIERRE=0 AND (CASITGUI='A' OR CASITGUI='V')"
rSTD01.Open SQL, cConexCom, adOpenStatic
'****Restaura Stock Actual
If Not rSTD01.EOF Then
'****Para ubicar Documento en El Almacen Destino
   AlmaRefe = rSTD01!CARFALMA
   NumRefe = rSTD01!CANUMDOC
End If
rSTD01.Close

SQL = " SELECT CAALMA,CARFALMA,CANUMDOC,CATD,CARFTDOC,CACODMOV,CAHORA,CAFECDOC " & _
     " FROM MovAlmCab WHERE CAALMA = '" & AlmaRefe & "'And CATD='NI' AND CACODMOV='TD' AND CARFNDOC ='" & NumRefe & "' AND CACIERRE=0 AND (CASITGUI='A' OR CASITGUI='V')"
rSTD01.Open SQL, cConexCom, adOpenStatic
If Not rSTD01.EOF Then
'****Para ubicar Documento en El Almacen Destino
   NumRefe = rSTD01!CANUMDOC
Else
   NumRefe = ""
End If
rSTD01.Close


'ANULO ORIGEN
ClsTDoc.AnuloCabezera AlmaActu, Tdoc, NumDoc, cConexCom
'ClsTDoc.EliminoDetalle AlmaActu, TDoc, NumDoc, cConexCom
'ANULO DESTINO
If AlmaRefe <> "" Then
   ClsTDoc.AnuloCabezera AlmaRefe, "NI", NumRefe, cConexCom
   'ClsTDoc.EliminoDetalle AlmaRefe, "NI", NumRefe, cConexCom
End If

For n = 1 To nCont
    SERLOTE = ClsTock.EsSerie_Lote(codigo(n), cConexCom)
    If SERLOTE = "S" Then
       ClsTock.CalculaStockSerie AlmaActu, codigo(n)
       If AlmaRefe <> "" Then ClsTock.CalculaStockSerie AlmaRefe, codigo(n)
    End If
    
    If SERLOTE = "L" Then
       ClsTock.CalculaStockLOTE AlmaActu, codigo(n)
       If AlmaRefe <> "" Then ClsTock.CalculaStockLOTE AlmaRefe, codigo(n)
   End If
   
    ClsTock.CalculaSaldoNoValorizado AlmaActu, codigo(n), Now   'Actualiza stkart , Moresmes
    If AlmaRefe <> "" Then ClsTock.CalculaSaldoNoValorizado AlmaRefe, codigo(n), Now
Next

    ClsTDoc.CorrigueNumeracion AlmaActu, cConexCom
    If AlmaRefe <> "" Then ClsTDoc.CorrigueNumeracion AlmaRefe, cConexCom
    
Exit Sub
ERRAR:
     MsgBox Err.Description
     
End Sub
Sub SetDatagrid()
End Sub
