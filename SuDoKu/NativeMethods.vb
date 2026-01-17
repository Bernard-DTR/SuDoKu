Imports System.Runtime.InteropServices

Class NativeMethods
  Declare Function GetTickCount64 Lib "kernel32" () As Integer
  Public Const MAX_ENTRY As Integer = 32768
  Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (
                     ByVal lpApplicationName As String,
                     ByVal lpKeyName As String,
                     ByVal lpDefault As String,
                     ByVal lpReturnedString As System.Text.StringBuilder,
                     ByVal nSize As Integer,
                     ByVal lpFileName As String) As Integer
  Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (
                     ByVal lpApplicationName As String,
                     ByVal lpKeyName As String,
                     ByVal lpString As String,
                     ByVal lpFileName As String) As Long

  <StructLayout(LayoutKind.Sequential)>
  Public Structure DEVMODE
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
    Public dmDeviceName As String
    Public dmSpecVersion As Short
    Public dmDriverVersion As Short
    Public dmSize As Short
    Public dmDriverExtra As Short
    Public dmFields As Integer
    Public dmPositionX As Integer
    Public dmPositionY As Integer
    Public dmDisplayOrientation As Integer
    Public dmDisplayFixedOutput As Integer
    Public dmColor As Short
    Public dmDuplex As Short
    Public dmYResolution As Short
    Public dmTTOption As Short
    Public dmCollate As Short
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
    Public dmFormName As String
    Public dmLogPixels As Short
    Public dmBitsPerPel As Short
    Public dmPelsWidth As Integer
    Public dmPelsHeight As Integer
    Public dmDisplayFlags As Integer
    Public dmDisplayFrequency As Integer
    Public dmICMMethod As Integer
    Public dmICMIntent As Integer
    Public dmMediaType As Integer
    Public dmDitherType As Integer
    Public dmReserved1 As Integer
    Public dmReserved2 As Integer
    Public dmPanningWidth As Integer
    Public dmPanningHeight As Integer
  End Structure

  <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
  Public Structure DISPLAY_DEVICE
    Public cb As Integer
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=32)>
    Public DeviceName As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
    Public DeviceString As String
    Public StateFlags As Integer
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
    Public DeviceID As String
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=128)>
    Public DeviceKey As String
  End Structure

  <DllImport("user32.dll")>
  Public Shared Function EnumDisplayDevices(deviceName As String, iDevNum As Integer, ByRef lpDisplayDevice As DISPLAY_DEVICE, dwFlags As Integer) As Boolean
  End Function

  <DllImport("user32.dll")>
  Public Shared Function EnumDisplaySettings(deviceName As String, modeNum As Integer, ByRef devMode As DEVMODE) As Boolean
  End Function

  '<DllImport("user32.dll")>
  Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
    Return True
  End Function

  Public Const ENUM_CURRENT_SETTINGS As Integer = -1
  Public Const EDD_GET_DEVICE_INTERFACE_NAME As Integer = &H1

End Class
