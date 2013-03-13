VERSION 5.00
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmImpresora 
   BackColor       =   &H80000009&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de impresora para la facturacion"
   ClientHeight    =   8730
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   Begin TrueOleDBGrid70.TDBGrid TDBG 
      Height          =   1695
      Left            =   120
      TabIndex        =   10
      Top             =   6840
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   2990
      _LayoutType     =   0
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).DataField=   ""
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).DataField=   ""
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   2
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   503
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   13160660
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=2"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(5)=   "Column(1).Width=2725"
      Splits(0)._ColumnProps(6)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(7)=   "Column(1)._WidthInPix=2646"
      Splits(0)._ColumnProps(8)=   "Column(1).Order=2"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=Tahoma"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      DefColWidth     =   0
      HeadLines       =   1
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   16777215
      RowDividerColor =   13160660
      RowSubDividerColor=   13160660
      DirectionAfterEnter=   1
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=Tahoma"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=Tahoma"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(11)  =   ":id=2,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=Tahoma"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bgcolor=&H8000000A&,.bold=0"
      _StyleDefs(14)  =   ":id=3,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=Tahoma"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Named:id=33:Normal"
      _StyleDefs(45)  =   ":id=33,.parent=0"
      _StyleDefs(46)  =   "Named:id=34:Heading"
      _StyleDefs(47)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(48)  =   ":id=34,.wraptext=-1"
      _StyleDefs(49)  =   "Named:id=35:Footing"
      _StyleDefs(50)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(51)  =   "Named:id=36:Selected"
      _StyleDefs(52)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(53)  =   "Named:id=37:Caption"
      _StyleDefs(54)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(55)  =   "Named:id=38:HighlightRow"
      _StyleDefs(56)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(57)  =   "Named:id=39:EvenRow"
      _StyleDefs(58)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(59)  =   "Named:id=40:OddRow"
      _StyleDefs(60)  =   ":id=40,.parent=33"
      _StyleDefs(61)  =   "Named:id=41:RecordSelector"
      _StyleDefs(62)  =   ":id=41,.parent=34"
      _StyleDefs(63)  =   "Named:id=42:FilterBar"
      _StyleDefs(64)  =   ":id=42,.parent=33"
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000009&
      Caption         =   "Tipo de Impresora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   135
      TabIndex        =   7
      Top             =   5355
      Width           =   3705
      Begin VB.OptionButton OptT 
         BackColor       =   &H80000009&
         Caption         =   "Ticketera"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2250
         TabIndex        =   9
         Top             =   270
         Width           =   1275
      End
      Begin VB.OptionButton OptM 
         BackColor       =   &H80000009&
         Caption         =   "Matricial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   495
         TabIndex        =   8
         Top             =   315
         Value           =   -1  'True
         Width           =   1500
      End
   End
   Begin VB.CommandButton CmdGrabar 
      BackColor       =   &H80000009&
      Caption         =   "&Grabar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5310
      Width           =   1185
   End
   Begin VB.CommandButton CmdRetornar 
      BackColor       =   &H80000009&
      Caption         =   "Retornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5310
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000009&
      Caption         =   "Establecer como Predeterminada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4140
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   135
      Width           =   3120
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   90
      TabIndex        =   0
      Top             =   585
      Width           =   7170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impresoras  Configuradas:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   6525
      Width           =   2235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Impresora Seleccionada :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   6
      Top             =   4275
      Width           =   2145
   End
   Begin VB.Label LblImpresora 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   600
      Left            =   90
      TabIndex        =   4
      Top             =   4545
      Width           =   7080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "IMPRESORAS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   270
      Width           =   1110
   End
End
Attribute VB_Name = "FrmImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function PtrCtoVbString(Add As Long) As String
    Dim sTemp As String * 512, X As Long

    X = lstrcpy(sTemp, Add)
    If (InStr(1, sTemp, Chr(0)) = 0) Then
         PtrCtoVbString = ""
    Else
         PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
    End If
End Function

Private Sub SetDefaultPrinter(ByVal PrinterName As String, _
    ByVal DriverName As String, ByVal PrinterPort As String)
    Dim DeviceLine As String
    Dim r As Long
    Dim l As Long
    DeviceLine = PrinterName & "," & DriverName & "," & PrinterPort
    ' Store the new printer information in the [WINDOWS] section of
    ' the WIN.INI file for the DEVICE= item
    r = WriteProfileString("windows", "Device", DeviceLine)
    ' Cause all applications to reload the INI file:
    l = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, "windows")
End Sub

Private Sub Win95SetDefaultPrinter()
    Dim Handle As Long          'handle to printer
    Dim PrinterName As String
    Dim pd As PRINTER_DEFAULTS
    Dim X As Long
    Dim need As Long            ' bytes needed
    Dim pi5 As PRINTER_INFO_5   ' your PRINTER_INFO structure
    Dim LastError As Long

    ' determine which printer was selected
    PrinterName = List1.List(List1.ListIndex)
    ' none - exit
    If PrinterName = "" Then
        Exit Sub
    End If

    ' set the PRINTER_DEFAULTS members
    pd.pDatatype = 0&
    pd.DesiredAccess = PRINTER_ALL_ACCESS Or pd.DesiredAccess

    ' Get a handle to the printer
    X = OpenPrinter(PrinterName, Handle, pd)
    ' failed the open
    If X = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' Make an initial call to GetPrinter, requesting Level 5
    ' (PRINTER_INFO_5) information, to determine how many bytes
    ' you need
    X = GetPrinter(Handle, 5, ByVal 0&, 0, need)
    ' don't want to check Err.LastDllError here - it's supposed
    ' to fail
    ' with a 122 - ERROR_INSUFFICIENT_BUFFER
    ' redim t as large as you need
    ReDim T((need \ 4)) As Long

    ' and call GetPrinter for keepers this time
    X = GetPrinter(Handle, 5, T(0), need, need)
    ' failed the GetPrinter
    If X = False Then
        'error handler code goes here
        Exit Sub
    End If

    ' set the members of the pi5 structure for use with SetPrinter.
    ' PtrCtoVbString copies the memory pointed at by the two string
    ' pointers contained in the t() array into a Visual Basic string.
    ' The other three elements are just DWORDS (long integers) and
    ' don't require any conversion
    pi5.pPrinterName = PtrCtoVbString(T(0))
    pi5.pPortName = PtrCtoVbString(T(1))
    pi5.Attributes = T(2)
    pi5.DeviceNotSelectedTimeout = T(3)
    pi5.TransmissionRetryTimeout = T(4)

    ' this is the critical flag that makes it the default printer
    pi5.Attributes = PRINTER_ATTRIBUTE_DEFAULT

       ' call SetPrinter to set it
       X = SetPrinter(Handle, 5, pi5, 0)

       If X = False Then   ' SetPrinter failed
           MsgBox "SetPrinter Failed. Error code: " & Err.LastDllError
           Exit Sub
       Else
           If Printer.DeviceName <> List1.Text Then
           ' Make sure Printer object is set to the new printer
                SelectPrinter (List1.Text)
           End If
       End If

    ' and close the handle
    ClosePrinter (Handle)
End Sub

Private Sub GetDriverAndPort(ByVal Buffer As String, DriverName As _
    String, PrinterPort As String)

    Dim iDriver As Integer
    Dim iPort As Integer
    DriverName = ""
    PrinterPort = ""

    ' The driver name is first in the string terminated by a comma
    iDriver = InStr(Buffer, ",")
    If iDriver > 0 Then

         ' Strip out the driver name
        DriverName = Left(Buffer, iDriver - 1)

        ' The port name is the second entry after the driver name
        ' separated by commas.
        iPort = InStr(iDriver + 1, Buffer, ",")

        If iPort > 0 Then
            ' Strip out the port name
            PrinterPort = Mid(Buffer, iDriver + 1, _
            iPort - iDriver - 1)
        End If
    End If
End Sub

Private Sub ParseList(lstCtl As Control, ByVal Buffer As String)
    Dim i As Integer
    Dim s As String

    Do
        i = InStr(Buffer, Chr(0))
        If i > 0 Then
            s = Left(Buffer, i - 1)
            If Len(Trim(s)) Then lstCtl.AddItem s
            Buffer = Mid(Buffer, i + 1)
        Else
            If Len(Trim(Buffer)) Then lstCtl.AddItem Buffer
            Buffer = ""
        End If
    Loop While i > 0
End Sub

Private Sub WinNTSetDefaultPrinter()
    Dim Buffer As String
    Dim DeviceName As String
    Dim DriverName As String
    Dim PrinterPort As String
    Dim PrinterName As String
    Dim r As Long
    If List1.ListIndex > -1 Then
        ' Get the printer information for the currently selected
        ' printer in the list. The information is taken from the
        ' WIN.INI file.
        Buffer = Space(1024)
        PrinterName = List1.Text
        r = GetProfileString("PrinterPorts", PrinterName, "", _
            Buffer, Len(Buffer))

        ' Parse the driver name and port name out of the buffer
        GetDriverAndPort Buffer, DriverName, PrinterPort

           If DriverName <> "" And PrinterPort <> "" Then
               SetDefaultPrinter List1.Text, DriverName, PrinterPort
               If Printer.DeviceName <> List1.Text Then
               ' Make sure Printer object is set to the new printer
                  SelectPrinter (List1.Text)
               End If
           End If
    End If
End Sub

Private Sub CmdGrabar_Click()
Dim rs As ADODB.Recordset
On Error GoTo Errores

If Len(Trim(LblImpresora.Caption)) = 0 Then
    MsgBox "Falta seleccionar impresora", vbInformation, "Configuracion"
    Exit Sub
End If

Set rs = VGCNx.Execute("select * from vt_configuraimpresora where " _
& " empresacodigo='" & VGParametros.empresacodigo & "' " _
& " and puntovtacodigo='" & VGParametros.puntovta & "' and " _
& " usuarionombre='" & VGParametros.cajerocodigo & "' and impresoratipo='" & IIf(OptM.Value = True, "M", "T") & "'")
If rs.RecordCount > 0 Then
    If MsgBox("La configuracion para este usuario ya existe:" & Chr(13) & "Desea modificarlo?", vbYesNo + vbQuestion, "Configuracion") = vbYes Then
        VGCNx.Execute "update vt_configuraimpresora set impresoranombre='" & LblImpresora.Caption & "' " _
        & " where empresacodigo='" & VGParametros.empresacodigo & "' " _
        & " and puntovtacodigo='" & VGParametros.puntovta & "' and " _
        & " usuarionombre='" & VGParametros.cajerocodigo & "' and impresoratipo='" & IIf(OptM.Value = True, "M", "T") & "'"
    
    MsgBox "Configuracion Modificada.!!!", vbInformation, "Sistema"
    
    End If
    
Else
    VGCNx.Execute "insert into vt_configuraimpresora(empresacodigo,puntovtacodigo," _
    & " usuarionombre,impresoranombre,impresoratipo) " _
    & " values ('" & VGParametros.empresacodigo & "','" & VGParametros.puntovta & "'," _
    & " '" & VGParametros.cajerocodigo & "','" & LblImpresora.Caption & "','" & IIf(OptM.Value = True, "M", "T") & "')"
    
    MsgBox "Configuracion grabada.", vbInformation, "Sistema"
    
End If

TDBG.Refresh

Unload Me

Exit Sub
Errores:
MsgBox "" & Err.Description & "", vbCritical, "Sistema"

End Sub

Private Sub CmdRetornar_Click()
Unload Me
End Sub

Private Sub Command1_Click()
Dim osinfo As OSVERSIONINFO
Dim retvalue As Integer
On Error GoTo Errores

osinfo.dwOSVersionInfoSize = 148
osinfo.szCSDVersion = Space$(128)
retvalue = GetVersionExA(osinfo)

If osinfo.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS Then
    Call Win95SetDefaultPrinter
Else
' This assumes that future versions of Windows use the NT method
    Call WinNTSetDefaultPrinter
    MsgBox "Impresora :" & Chr(13) & List1.Text & Chr(13) & "Establecida como predeterminada.", vbInformation, "Sistema"
End If



Exit Sub
Errores:
MsgBox ""

End Sub

Private Sub Command2_Click()

End Sub


Private Sub Form_Load()
Dim r As Long
Dim Buffer As String
Dim RsImpresoras As ADODB.Recordset

Me.Left = 100
Me.Top = 100

' Get the list of available printers from WIN.INI
Buffer = Space(8192)
r = GetProfileString("PrinterPorts", vbNullString, "", _
Buffer, Len(Buffer))

' Display the list of printer in the ListBox List1
ParseList List1, Buffer

Set RsImpresoras = VGCNx.Execute("select impresoranombre as Impresora,Tipo=case when impresoratipo='M' then 'MATRICIAL' else 'TICKETERA' end from vt_configuraimpresora where " _
& " empresacodigo='" & VGParametros.empresacodigo & "' and puntovtacodigo='" & VGParametros.puntovta & "' order by 2")
If RsImpresoras.RecordCount > 0 Then
    TDBG.DataSource = RsImpresoras
    TDBG.Refresh
End If

TDBG.Columns(0).Width = 5600
TDBG.Columns(1).Width = 1100


CmdGrabar.Picture = MDIPrincipal.ImageList2.ListImages.Item("Imprimir").Picture
CmdRetornar.Picture = MDIPrincipal.ImageList2.ListImages.Item("Retornar").Picture

End Sub

Private Sub List1_Click()
LblImpresora.Caption = List1.Text
NombreImpresora = LblImpresora.Caption
End Sub
