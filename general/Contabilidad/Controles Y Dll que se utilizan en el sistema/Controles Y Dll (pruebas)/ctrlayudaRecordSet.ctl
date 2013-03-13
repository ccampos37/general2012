VERSION 5.00
Begin VB.UserControl CtrAy_RS 
   ClientHeight    =   348
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2124
   LockControls    =   -1  'True
   ScaleHeight     =   348
   ScaleWidth      =   2124
   Begin VB.CommandButton CmdEjecutarAyuda 
      Caption         =   "..."
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   15
      Width           =   255
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000000&
      Height          =   285
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   15
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   285
   End
End
Attribute VB_Name = "CtrAy_RS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Default Property Values:
Const m_def_ListaCamposText = ""
Const m_def_ListaCamposDescrip = ""
Const m_def_TituloAyuda = ""
Const m_def_ListaCampos = ""
Const m_def_NomTabla = ""
Const m_def_xcodwith = 0
Const m_def_Enabled = 0
Const m_def_XcodCampo = ""
Const m_def_XListCampo = ""
Const m_def_xcodigomaxleng = 0
'Property Variables:
Dim m_ListaCamposText As Variant
Dim m_ListaCamposDescrip As Variant
Dim m_TituloAyuda As String
Dim m_ListaCampos As Variant
Dim m_NomTabla As String
Dim m_xcodwith As Long
Dim m_Enabled As Boolean
Dim m_XcodCampo As String
Dim m_XListCampo As String
Dim m_filtro As Variant
Dim m_xcodigomaxleng As Long
'variables privadas
Dim v_anchotext1 As Long
Dim v_anchotext2 As Long
Dim v_cont  As Long, v_lefttext2 As Long
Dim v_primercampo As String
Dim rs_consul As ADODB.Recordset
Dim cnx As ADODB.Connection
Dim ayudadll As New ClassFormAyuda
Dim m_seconecto As Boolean
Dim m_varauxtext As String
Dim m_presionetab As Boolean
Dim m_presioneEnter As Boolean
'Event Declarations:
Event Click()
Attribute Click.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse encima de un objeto."
Event DblClick()
Attribute DblClick.VB_Description = "Ocurre cuando el usuario presiona y libera un botón del mouse y después lo vuelve a presionar y liberar sobre un objeto."
Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Ocurre cuando el usuario presiona una tecla mientras un objeto tiene el enfoque."
Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Ocurre cuando el usuario presiona y libera una tecla ANSI."
Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Ocurre cuando el usuario libera una tecla mientras un objeto tiene el enfoque."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Ocurre cuando el usuario presiona el botón del mouse mientras un objeto tiene el enfoque."
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Ocurre cuando el usuario mueve el mouse."
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Ocurre cuando el usuario libera el botón del mouse mientras un objeto tiene el enfoque."
Public Event AlDevolverDato(ByVal ColecCampos As Fields)
Attribute AlDevolverDato.VB_MemberFlags = "200"
Public Event AlNoDevolverNada()
Private Sub CmdEjecutarAyuda_Click()
    If Not m_seconecto Then
        MsgBox "Tiene que Utilizar el Metodo => ¡Conexion! de Este Control", vbCritical
        Exit Sub
    End If

Dim camp As Fields, nreg As Long
    Set camp = Nothing
    ayudadll.tabla = NomTabla
    ayudadll.ListaCampos = ListaCampos
    ayudadll.TituloAyuda = TituloAyuda
    ayudadll.PrimerCampo = XcodCampo
    ayudadll.ListaCamposDescrip = ListaCamposDescrip
    ayudadll.Filtro = m_filtro
    Call ayudadll.mostrar(cnx, camp, nreg)
    If Not camp Is Nothing And nreg > 0 Then
        xclave = camp(XcodCampo).Value
        xnombre = camp(XListCampo).Value
        RaiseEvent AlDevolverDato(camp)
        Call SendKeys("{TAB}")
     Else
        RaiseEvent AlNoDevolverNada
    End If
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Devuelve o establece un valor que determina si un objeto puede responder a eventos generados por el usuario."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Text1,Text1,-1,Text
Public Property Get xclave() As String
Attribute xclave.VB_Description = "Devuelve o Establece el contenido del campo clave"
    xclave = Text1.Text
End Property

Public Property Let xclave(ByVal New_xclave As String)
    Text1.Text() = New_xclave
    PropertyChanged "xclave"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MappingInfo=Text2,Text2,-1,Text
Public Property Get xnombre() As String
Attribute xnombre.VB_Description = "Devuelve el texto del nombre"
    xnombre = Text2.Text
End Property

Public Property Let xnombre(ByVal New_xnombre As String)
    Text2.Text() = New_xnombre
    PropertyChanged "xnombre"
End Property

Private Sub Text1_GotFocus()
    Text1.BackColor = &HCFFBFC
    m_presionetab = True
    m_varauxtext = Trim(Text1.Text)
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 144 And Shift = 0 Then
        m_presionetab = True
      Else
        m_presionetab = False
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If Chr(KeyAscii) = "'" Then KeyAscii = 0

    If KeyAscii = 13 Then
        m_presioneEnter = True
        Call SendKeys("{TAB}")
      Else
        m_presioneEnter = False
    End If
End Sub
Private Sub ConsultarenText()
Dim lista As Variant
    If Not m_seconecto Then
        MsgBox "Tiene que Utilizar el Metodo => ¡Conexion! de Este Control", vbCritical
        Exit Sub
    End If
    If Trim(ListaCamposText) = "" Then
        MsgBox "No se ha especificado la lista de campos en la Propiedad=>ListaCamposText", vbCritical, "Componente"
        Exit Sub
    End If
    If Not ayudadll.ExisteTablaoCampos(NomTabla, cnx, Trim(ListaCamposText)) Then
        MsgBox "No se ha tipeado bien la Lista de campos en la Propiedad=>ListaCamposText", vbCritical, "Componente"
        Exit Sub
    End If
    
    If (Trim(m_varauxtext) <> Trim(Text1.Text)) And Trim(Text1.Text) <> "" Then
        m_varauxtext = Trim(Text1.Text)
      Else
        If Trim(Text1.Text) <> "" And m_presioneEnter Then
            Call SendKeys("{TAB}{TAB}")
        End If
        If Trim(Text1.Text) = "" Then
            Text2.Text = ""
        End If
        Exit Sub
    End If
    
    Set rs_consul = New ADODB.Recordset
    rs_consul.Open "select " & ListaCamposText & " from " & NomTabla & _
                   " where " & XcodCampo & "='" & Trim(Text1.Text) & "'" & IIf(Trim(m_filtro) = "", "", " and (" & m_filtro & ")"), cnx, adOpenKeyset, adLockReadOnly
    If rs_consul.RecordCount = 0 Or rs_consul.RecordCount = -1 Then
        xclave = "": xnombre = ""
        RaiseEvent AlNoDevolverNada
     Else
       xclave = rs_consul.Fields(XcodCampo): xnombre = rs_consul.Fields(XListCampo)
       RaiseEvent AlDevolverDato(rs_consul.Fields)
       Call SendKeys("{TAB}{TAB}")
    End If
End Sub
Private Sub Text1_LostFocus()
    Call ConsultarenText
    Text1.BackColor = vbWhite
    m_presioneEnter = False
End Sub

Private Sub Text2_GotFocus()
    If Text1.Text = "" And Not m_presionetab Then Call SendKeys("+{TAB}")
    If Text1.Text = "" And m_presionetab Then Call SendKeys("{TAB}")
End Sub
Private Sub UserControl_Initialize()
    v_anchotext1 = Text1.Width
    v_anchotext2 = Text2.Width
    v_cont = 0
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
    m_xcodwith = m_def_xcodwith
    m_NomTabla = m_def_NomTabla
    m_TituloAyuda = m_def_TituloAyuda
    m_ListaCampos = m_def_ListaCampos
    m_XcodCampo = m_def_XcodCampo
    m_XListCampo = m_def_XListCampo
    m_ListaCamposDescrip = m_def_ListaCamposDescrip
    m_ListaCamposText = m_def_ListaCamposText
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    Text1.Text = PropBag.ReadProperty("xclave", "")
    Text2.Text = PropBag.ReadProperty("xnombre", "")
    Text1.MaxLength = PropBag.ReadProperty("XcodMaxLongitud", 0)
    m_xcodwith = PropBag.ReadProperty("xcodwith", m_def_xcodwith)
    Call dimensionatext1(xcodwith)
    m_NomTabla = PropBag.ReadProperty("NomTabla", m_def_NomTabla)
    m_TituloAyuda = PropBag.ReadProperty("TituloAyuda", m_def_TituloAyuda)
    m_ListaCampos = PropBag.ReadProperty("ListaCampos", m_def_ListaCampos)
    m_XcodCampo = PropBag.ReadProperty("XcodCampo", m_def_XcodCampo)
    m_XListCampo = PropBag.ReadProperty("XListCampo", m_def_XListCampo)
    m_ListaCamposDescrip = PropBag.ReadProperty("ListaCamposDescrip", m_def_ListaCamposDescrip)
    m_ListaCamposText = PropBag.ReadProperty("ListaCamposText", m_def_ListaCamposText)
End Sub

Private Sub UserControl_Resize()
    Text2.Width = Text2.Width + (Width - (Text1.Width + Text2.Width + 290))
    CmdEjecutarAyuda.Left = Text2.Left + Text2.Width + 20
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("xclave", Text1.Text, "")
    Call PropBag.WriteProperty("xnombre", Text2.Text, "")
    Call PropBag.WriteProperty("XcodMaxLongitud", Text1.MaxLength, "")
    Call PropBag.WriteProperty("xcodwith", m_xcodwith, m_def_xcodwith)
    Call PropBag.WriteProperty("NomTabla", m_NomTabla, m_def_NomTabla)
    Call PropBag.WriteProperty("TituloAyuda", m_TituloAyuda, m_def_TituloAyuda)
    Call PropBag.WriteProperty("ListaCampos", m_ListaCampos, m_def_ListaCampos)
    Call PropBag.WriteProperty("XcodCampo", m_XcodCampo, m_def_XcodCampo)
    Call PropBag.WriteProperty("XListCampo", m_XListCampo, m_def_XListCampo)
    Call PropBag.WriteProperty("ListaCamposDescrip", m_ListaCamposDescrip, m_def_ListaCamposDescrip)
    Call PropBag.WriteProperty("ListaCamposText", m_ListaCamposText, m_def_ListaCamposText)
End Sub

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=8,0,0,0
Public Property Get xcodwith() As Long
Attribute xcodwith.VB_Description = "el ancho del campo codigo"
    xcodwith = m_xcodwith
End Property

Public Property Let xcodwith(ByVal New_xcodwith As Long)
    Call dimensionatext1(New_xcodwith)
    m_xcodwith = New_xcodwith
    PropertyChanged "xcodwith"
End Property

Private Sub dimensionatext1(valor As Long)
    Text1.Width = v_anchotext1 + valor
    If v_cont = 0 Then v_lefttext2 = Text2.Left
    v_cont = v_cont + 1
    Text2.Width = Width - (Text1.Width + CmdEjecutarAyuda.Width + 30)
    Text2.Left = v_lefttext2 + (Text1.Width - v_anchotext1)
End Sub

Public Sub conexion(ByVal conex As ADODB.Connection)
    If conex Is Nothing Then
        m_seconecto = False
        MsgBox "La Conexion se Encuentra Cerra o No Establecida => En el " & _
               "Metodo Conexion de Este Control ", vbCritical
        Exit Sub
      Else
        m_seconecto = True
        m_varauxtext = "''"
        m_presionetab = True
    End If
    Set cnx = conex
End Sub
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,0
Public Property Get NomTabla() As String
Attribute NomTabla.VB_Description = "Se Establece el Nombre de la Tabla de Ayuda o de Busqueda"
    NomTabla = m_NomTabla
End Property

Public Property Let NomTabla(ByVal New_NomTabla As String)
    m_NomTabla = New_NomTabla
    PropertyChanged "NomTabla"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=13,0,0,0
Public Property Get TituloAyuda() As String
Attribute TituloAyuda.VB_Description = "Se establece el titulo que aparecera en la ventana de ayuda"
    TituloAyuda = m_TituloAyuda
End Property

Public Property Let TituloAyuda(ByVal New_TituloAyuda As String)
    m_TituloAyuda = New_TituloAyuda
    PropertyChanged "TituloAyuda"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get ListaCampos() As Variant
Attribute ListaCampos.VB_Description = "Aqui Establece la lista de campos para la busqueda en la ventana de Ayuda"
    ListaCampos = m_ListaCampos
End Property

Public Property Let ListaCampos(ByVal New_ListaCampos As Variant)
    m_ListaCampos = New_ListaCampos
    PropertyChanged "ListaCampos"
End Property
Public Property Get XcodCampo() As String
Attribute XcodCampo.VB_Description = "Se establece el  campo de busqueda principal"
    XcodCampo = m_XcodCampo
End Property

Public Property Let XcodCampo(ByVal New_XcodCampo As String)
    m_XcodCampo = New_XcodCampo
    PropertyChanged "XcodCampo"
End Property
Public Property Get XListCampo() As String
Attribute XListCampo.VB_Description = "El campo que se mostrara en el campo nombre"
    XListCampo = m_XListCampo
End Property

Public Property Let XListCampo(ByVal New_XListCampo As String)
    m_XListCampo = New_XListCampo
    PropertyChanged "XListCampo"
End Property
'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get ListaCamposDescrip() As Variant
    ListaCamposDescrip = m_ListaCamposDescrip
End Property

Public Property Let ListaCamposDescrip(ByVal New_ListaCamposDescrip As Variant)
    m_ListaCamposDescrip = New_ListaCamposDescrip
    PropertyChanged "ListaCamposDescrip"
End Property

'ADVERTENCIA: NO QUITAR NI MODIFICAR LAS SIGUIENTES LINEAS CON COMENTARIOS
'MemberInfo=14,0,0,0
Public Property Get ListaCamposText() As Variant
    ListaCamposText = m_ListaCamposText
End Property
Public Property Let ListaCamposText(ByVal New_ListaCamposText As Variant)
    m_ListaCamposText = New_ListaCamposText
    PropertyChanged "ListaCamposText"
End Property
Public Property Let Filtro(ByVal New_valor As Variant)
    m_filtro = New_valor
End Property
Public Property Get XcodMaxLongitud() As Long
Attribute XcodMaxLongitud.VB_Description = "Longitud del text de Codigo del Campo o El Primer\r\nText"
    XcodMaxLongitud = Text1.MaxLength
End Property
Public Property Let XcodMaxLongitud(ByVal New_XcodMaxLongitud As Long)
    Text1.MaxLength() = New_XcodMaxLongitud
    PropertyChanged "XcodMaxLongitud"
End Property


