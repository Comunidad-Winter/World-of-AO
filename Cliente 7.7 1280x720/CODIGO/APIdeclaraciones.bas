Attribute VB_Name = "Api"
Public Declare Function FindWindow _
                         Lib "user32" _
                             Alias "FindWindowA" (ByVal lpClassName As String, _
                                                  ByVal lpWindowName As String) As Long

Public Const WM_SETTEXT = &HC
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const EM_SETREADONLY = &HCF

Public Declare Function EnumDisplaySettings _
                         Lib "user32" _
                             Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, _
                                                           ByVal iModeNum As Long, _
                                                           lptypDevMode As Any) As Boolean
Public Declare Function ChangeDisplaySettings _
                         Lib "user32" _
                             Alias "ChangeDisplaySettingsA" (lptypDevMode As Any, _
                                                             ByVal dwFlags As Long) As Long

'[MatuX] : 24 de Marzo del 2002
Declare Function SetWindowPos& _
                  Lib "user32" (ByVal hwnd As Long, _
                                ByVal hWndInsertAfter As Long, _
                                ByVal X As Long, _
                                ByVal Y As Long, _
                                ByVal cx As Long, _
                                ByVal cy As Long, _
                                ByVal wFlags As Long)
'[END]

Public Const CCDEVICENAME = 32
Public Const CCFORMNAME = 32
Public Const DM_BITSPERPEL = &H40000
Public Const DM_PELSWIDTH = &H80000
Public Const DM_PELSHEIGHT = &H100000
Public Const CDS_UPDATEREGISTRY = &H1
Public Const CDS_TEST = &H4
Public Const DISP_CHANGE_SUCCESSFUL = 0
Public Const DISP_CHANGE_RESTART = 1

Type typDevMODE

    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long

End Type
