Imports System.Runtime.InteropServices
Imports Microsoft.Win32

Module Program

    Sub Main(CmdArgs() As String)
        Console.WriteLine("VbLoadDriver" & Environment.NewLine & "http://vbscrub.com" & Environment.NewLine)
        Try
            If CmdArgs.Count < 1 OrElse CmdArgs.Count > 2 OrElse CmdArgs(0) = "/?" Then
                DisplayUsage()
                Exit Sub
            End If

            Dim OriginalRegPath As String = CmdArgs(0)
            Dim DriverPath As String = Nothing
            Dim KernelRegPath As String = Nothing
            Dim Luid As WinApi.LUID
            Dim TokenHandle As IntPtr
            Dim IsUserKey As Boolean

            If CmdArgs.Count = 2 Then
                DriverPath = CmdArgs(1)
            End If

            If OriginalRegPath.StartsWith("HKLM", StringComparison.CurrentCultureIgnoreCase) Then
                KernelRegPath = "\Registry\Machine" & OriginalRegPath.Substring(4)
                IsUserKey = False
            ElseIf OriginalRegPath.StartsWith("HKU", StringComparison.CurrentCultureIgnoreCase) Then
                KernelRegPath = "\Registry\User" & OriginalRegPath.Substring(3)
                IsUserKey = True
            Else
                Console.WriteLine("Invalid registry path specified. Must begin with HKLM or HKU")
                Exit Sub
            End If

            Console.WriteLine("Attempting to enable SeLoadDriverPrivilege...")
            If Not WinApi.LookupPrivilegeValue(Nothing, WinApi.SE_LOAD_DRIVER_NAME, Luid) Then
                Console.WriteLine("Failed to get LUID for SeLoadDriverPrivilege : " & New ComponentModel.Win32Exception().Message)
                Exit Sub
            End If

            Try
                Using CurrentProcess As Process = Process.GetCurrentProcess
                    If Not WinApi.OpenProcessToken(CurrentProcess.Handle, WinApi.TOKEN_QUERY Or WinApi.TOKEN_ADJUST_PRIVILEGES, TokenHandle) Then
                        Console.WriteLine("Failed to open current process token: " & New ComponentModel.Win32Exception().Message)
                        Exit Sub
                    End If
                End Using

                WinApi.AdjustTokenPrivileges(TokenHandle, False, GetPriv(Luid), Marshal.SizeOf(GetType(WinApi.TOKEN_PRIVILEGES)), Nothing, Nothing)
                Dim LastError As Integer = Marshal.GetLastWin32Error
                If Not LastError = 0 Then
                    Console.WriteLine("Failed to enable SeLoadDriverPrivilege: " & New ComponentModel.Win32Exception(LastError).Message)
                    Exit Sub
                End If
                Console.WriteLine("Successfully enabled privilege")

                If Not String.IsNullOrEmpty(DriverPath) Then
                    Console.WriteLine("Creating registry values...")
                    Dim RootReg As RegistryKey = Nothing
                    If IsUserKey Then
                        RootReg = RegistryKey.OpenBaseKey(RegistryHive.Users, RegistryView.Default)
                    Else
                        RootReg = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Default)
                    End If
                    Using RootReg
                        Using Subkey As RegistryKey = RootReg.CreateSubKey(OriginalRegPath.Substring(OriginalRegPath.IndexOf("\"c) + 1))
                            Subkey.SetValue("ImagePath", "\??\" & DriverPath, RegistryValueKind.ExpandString)
                            Subkey.SetValue("Type", 1, RegistryValueKind.DWord)
                            Subkey.SetValue("ErrorControl", 1, RegistryValueKind.DWord)
                            Subkey.SetValue("Start", 3, RegistryValueKind.DWord)
                        End Using
                    End Using
                    Console.WriteLine("Successfully created registry values")
                End If

                Console.WriteLine("Loading driver...")
                Dim NtResult As UInteger = WinApi.NtLoadDriver(GetUnicodeString(KernelRegPath))
                If NtResult > WinApi.NT_ERROR_START Then
                    Console.WriteLine("NtLoadDriver returned error code 0x" & Hex(NtResult))
                    Console.WriteLine(New ComponentModel.Win32Exception(WinApi.RtlNtStatusToDosError(NtResult)).Message)
                    Exit Sub
                End If
                Console.WriteLine("Successfully loaded driver")
            Finally
                If Not TokenHandle = IntPtr.Zero Then
                    WinApi.CloseHandle(TokenHandle)
                End If
            End Try
        Catch ex As Exception
            Console.WriteLine("Error: " & ex.Message)
        End Try

    End Sub

    Private Function GetPriv(Luid As WinApi.LUID) As WinApi.TOKEN_PRIVILEGES
        Dim Privs As WinApi.TOKEN_PRIVILEGES
        Dim PrivAttributes(0) As WinApi.LUID_AND_ATTRIBUTES
        PrivAttributes(0).Luid = Luid
        PrivAttributes(0).Attributes = WinApi.SE_PRIVILEGE_ENABLED
        Privs.Privileges = PrivAttributes
        Privs.PrivilegeCount = 1
        Return Privs
    End Function

    Private Function GetUnicodeString(Source As String) As WinApi.UNICODE_STRING
        Dim UnicodeString As WinApi.UNICODE_STRING
        UnicodeString.Buffer = Source
        UnicodeString.Length = CUShort(UnicodeString.Buffer.Length * 2)
        UnicodeString.MaximumLength = CUShort(UnicodeString.Length + 2)
        Return UnicodeString
    End Function

    Private Sub DisplayUsage()
        Console.WriteLine("Enables the driver loading priviledge and then loads a specified driver from" & Environment.NewLine & "the registry" & Environment.NewLine & Environment.NewLine &
                          "USAGE: VbLoadDriver.exe RegistryPath [DriverPath]" & Environment.NewLine & Environment.NewLine &
                           "RegistryPath" & vbTab & "" & "Path to registry key that contains parameters for this driver." & Environment.NewLine &
                           vbTab & vbTab & "Must start with HKLM (for HKEY_LOCAL_MACHINE) or HKU (for" & Environment.NewLine &
                           vbTab & vbTab & "HKEY_USERS). If loading a new driver, specify a non-existent" & Environment.NewLine &
                           vbTab & vbTab & "registry key here and the program will create the registry key" & Environment.NewLine &
                           vbTab & vbTab & "and populate it with values that use the driver file specified" & Environment.NewLine &
                           vbTab & vbTab & "by the DriverPath argument." & Environment.NewLine & Environment.NewLine &
                           "DriverPath" & vbTab & "Full path to the driver file to be loaded, which will be" & Environment.NewLine &
                           vbTab & vbTab & "written to the registry key specified by the RegistryPath" & Environment.NewLine &
                           vbTab & vbTab & "argument. This argument is optional. Using it when specifying" & Environment.NewLine &
                           vbTab & vbTab & "an existing registry key for RegistryPath will cause new values" & Environment.NewLine &
                           vbTab & vbTab & "to be written to the registry key, overwriting existing values." & Environment.NewLine & Environment.NewLine &
                           "EXAMPLES: " & Environment.NewLine & Environment.NewLine &
                          "VbLoadDriver.exe HKLM\System\CurrentControlSet\Services\SomeDriver" & Environment.NewLine & Environment.NewLine &
                          "VbLoadDriver.exe HKU\S-1-5-21-1950...58-1040\System\CurrentControlSet\Services\MyNewDriver C:\NewDriver.sys" & Environment.NewLine)

    End Sub

End Module
