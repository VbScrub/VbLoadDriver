Imports System.Runtime.InteropServices

Public Class WinApi

    Public Const SE_PRIVILEGE_ENABLED As Integer = 2
    Public Const SE_LOAD_DRIVER_NAME As String = "SeLoadDriverPrivilege"
    Public Const NT_ERROR_START As UInteger = 3221225472 '0xC0000000
    Public Const TOKEN_QUERY As Integer = 8
    Public Const TOKEN_ADJUST_PRIVILEGES As Integer = 32 '0x20

    <StructLayoutAttribute(LayoutKind.Sequential)>
    Public Structure UNICODE_STRING
        Public Length As UShort
        Public MaximumLength As UShort
        <MarshalAsAttribute(UnmanagedType.LPWStr)>
        Public Buffer As String
    End Structure

    <StructLayoutAttribute(LayoutKind.Sequential)>
        Public Structure TOKEN_PRIVILEGES
        Public PrivilegeCount As UInteger
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=1, ArraySubType:=UnmanagedType.Struct)>
        Public Privileges() As LUID_AND_ATTRIBUTES
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure LUID
        Public LowPart As UInteger
        Public HighPart As Integer
    End Structure

    <StructLayoutAttribute(LayoutKind.Sequential)>
    Public Structure LUID_AND_ATTRIBUTES
        Public Luid As LUID
        Public Attributes As UInteger
    End Structure

    <DllImport("ntdll.dll", SetLastError:=True)>
    Public Shared Function RtlNtStatusToDosError(NtStatus As UInteger) As Integer
    End Function

    <DllImport("ntdll.dll", SetLastError:=True)>
    Public Shared Function NtLoadDriver(ByRef DriverServiceName As UNICODE_STRING) As UInteger
    End Function

    <DllImportAttribute("advapi32.dll", EntryPoint:="LookupPrivilegeValueW", SetLastError:=True)>
    Public Shared Function LookupPrivilegeValue(<MarshalAs(UnmanagedType.LPWStr)> ByVal lpSystemName As String,
                                                <MarshalAs(UnmanagedType.LPWStr)> ByVal lpName As String,
                                                <Out()> ByRef lpLuid As LUID) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", SetLastError:=True)>
    Public Shared Function AdjustTokenPrivileges(ByVal TokenHandle As IntPtr,
                                                 <MarshalAs(UnmanagedType.Bool)> ByVal DisableAllPrivileges As Boolean,
                                                 ByRef NewState As TOKEN_PRIVILEGES,
                                                 ByVal BufferLength As Integer,
                                                 <Out()> ByVal PreviousState As IntPtr,
                                                 <Out()> ByRef ReturnLength As UInteger) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("advapi32.dll", EntryPoint:="OpenProcessToken", SetLastError:=True)>
    Public Shared Function OpenProcessToken(ByVal ProcessHandle As IntPtr,
                                            ByVal DesiredAccess As UInteger,
                                            ByRef TokenHandle As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)> _
    Public Shared Function CloseHandle(ByVal hObject As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function




End Class
