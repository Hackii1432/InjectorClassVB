Imports System
Imports System.Diagnostics
Imports System.Runtime.InteropServices
Public Class InjectorClass
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function OpenProcess(ByVal dwDesiredAccess As UInt32, ByVal bInheritHandle As Int32, ByVal dwProcessId As UInt32) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function CloseHandle(ByVal hObject As IntPtr) As Int32
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function GetProcAddress(ByVal hModule As IntPtr, ByVal lpProcName As String) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function VirtualAllocEx(ByVal hProcess As IntPtr, ByVal lpAddress As IntPtr, ByVal dwSize As IntPtr, ByVal flAllocationType As UInteger, ByVal flProtect As UInteger) As IntPtr
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function WriteProcessMemory(ByVal hProcess As IntPtr, ByVal lpBaseAddress As IntPtr, ByVal buffer As Byte(), ByVal size As UInteger, ByRef lpNumberOfBytesWritten As IntPtr) As Int32
    End Function
    <DllImport("kernel32.dll", SetLastError:=True)> _
    Private Shared Function CreateRemoteThread(ByVal hProcess As IntPtr, ByVal lpThreadAttribute As IntPtr, ByVal dwStackSize As IntPtr, ByVal lpStartAddress As IntPtr, ByVal lpParameter As IntPtr, ByVal dwCreationFlags As UInteger, _
     ByVal lpThreadId As IntPtr) As IntPtr
    End Function
    Public Class VAE_Enums
        Public Enum AllocationType
            MEM_COMMIT = &H1000
            MEM_RESERVE = &H2000
            MEM_RESET = &H80000
        End Enum
        Public Enum ProtectionConstants
            PAGE_EXECUTE = &H10
            PAGE_EXECUTE_READ = &H20
            PAGE_EXECUTE_READWRITE = &H40
            PAGE_EXECUTE_WRITECOPY = &H80
            PAGE_NOACCESS = &H1
        End Enum
    End Class
    Public Shared Function DoInject(ByVal pToBeInjected As Process, ByVal sDllPath As String, ByRef sError As String) As Boolean
        Dim hwnd As IntPtr = IntPtr.Zero
        If Not CRT(pToBeInjected, sDllPath, sError, hwnd) Then
            'CreateRemoteThread
            'close the handle, since the method wasn't able to get to that
            If hwnd <> New IntPtr(0) Then
                CloseHandle(hwnd)
            End If
            Return False
        End If
        Dim wee As Integer = Marshal.GetLastWin32Error()
        Return True
    End Function
    Private Shared Function CRT(ByVal pToBeInjected As Process, ByVal sDllPath As String, ByRef sError As String, ByRef hwnd As IntPtr) As Boolean
        sError = [String].Empty
        'in case we encounter no errors
        'create thread, query info, operation
        'write, and read
        Dim hndProc As IntPtr = OpenProcess((&H2 Or &H8 Or &H10 Or &H20 Or &H400), 1, CUInt(pToBeInjected.Id))

        hwnd = hndProc

        If hndProc = New IntPtr(0) Then
            sError = "Unable to attatch to process." & vbLf
            sError += "Error code: " & Marshal.GetLastWin32Error()
            Return False
        End If

        Dim lpLLAddress As IntPtr = GetProcAddress(GetModuleHandle("kernel32.dll"), "LoadLibraryA")

        If lpLLAddress = New IntPtr(0) Then
            sError = "Unable to find address of ""LoadLibraryA""." & vbLf
            sError += "Error code: " & Marshal.GetLastWin32Error()
            Return False
        End If

        '520 bytes should be enough
        Dim lpAddress As IntPtr = VirtualAllocEx(hndProc, New IntPtr(0), New IntPtr(sDllPath.Length), CUInt(VAE_Enums.AllocationType.MEM_COMMIT) Or CUInt(VAE_Enums.AllocationType.MEM_RESERVE), CUInt(VAE_Enums.ProtectionConstants.PAGE_EXECUTE_READWRITE))

        If lpAddress = New IntPtr(0) Then
            If lpAddress = New IntPtr(0) Then
                sError = "Unable to allocate memory to target process." & vbLf
                sError += "Error code: " & Marshal.GetLastWin32Error()
                Return False
            End If
        End If

        Dim bytes As Byte() = CalcBytes(sDllPath)
        Dim ipTmp As IntPtr = IntPtr.Zero

        WriteProcessMemory(hndProc, lpAddress, bytes, CUInt(bytes.Length), ipTmp)

        If Marshal.GetLastWin32Error() <> 0 Then
            sError = "Unable to write memory to process."
            sError += "Error code: " & Marshal.GetLastWin32Error()
            Return False
        End If

        Dim ipThread As IntPtr = CreateRemoteThread(hndProc, New IntPtr(0), New IntPtr(0), lpLLAddress, lpAddress, 0, _
         New IntPtr(0))

        If ipThread = New IntPtr(0) Then
            sError = "Unable to load dll into memory."
            sError += "Error code: " & Marshal.GetLastWin32Error()
            Return False
        End If
        Return True
    End Function
    Private Shared Function CalcBytes(ByVal sToConvert As String) As Byte()
        Dim bRet As Byte() = System.Text.Encoding.ASCII.GetBytes(sToConvert)
        Return bRet
    End Function
End Class
