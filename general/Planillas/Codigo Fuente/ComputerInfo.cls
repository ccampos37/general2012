VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComputerInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lbSectorPerCluster As Long, lbBytesPerSector As Long, lbNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public Property Get UserName()
    Dim sName As String
    Dim iRetVal As Long
    sName = Space$(255)
    iRetVal = GetUserName(sName, 255&)
    If iRetVal = 0 Then Exit Property
    UserName = sName
End Property

Private Function FreeSpace(unidad As String) As Double
Dim sDrive As String
Dim lreturn As Long
Dim l1 As Long 'Sectores por cluster
Dim l2 As Long 'bytes por sector
Dim l3 As Long 'numero de cluster libres
Dim l4 As Long 'total de cluster

sDrive = Left$(unidad, 1) & ":\"
lreturn = GetDiskFreeSpace(sDrive, l1, l2, l3, l4)

FreeSpace = (l1 * l2 * l3)

End Function

Public Property Get ComputerName() As Variant
    Dim sName As String
    Dim iRetVal As Long
    Dim ipos As Integer
    
    sName = Space$(255)
    iRetVal = GetComputerName(sName, 255&)
    If iRetVal = 0 Then
      ComputerName = ""
      Exit Property
    End If
    ipos = InStr(sName, Chr$(0))
    ComputerName = Left$(sName, ipos - 1)
End Property

Public Property Get FreeSpaceBytes(unidad As String) As Double
FreeSpaceBytes = FreeSpace(unidad)

End Property

Public Property Get FreeSpaceKbBytes(unidad As String) As Double
FreeSpaceKbBytes = Round(FreeSpace(unidad) / 1024, 2)
End Property

Public Property Get FreeSpaceMbBytes(unidad As String) As Double
FreeSpaceMbBytes = Round(FreeSpace(unidad) / 1048576, 2)
End Property

Public Property Get FreeSpaceGbBytes(unidad As String) As Double
FreeSpaceGbBytes = Round(FreeSpace(unidad) / 1073741824, 2)
End Property
