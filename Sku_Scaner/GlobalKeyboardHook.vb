Imports System.Runtime.InteropServices
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Module GlobalKeyboardHook
    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WM_KEYDOWN As Integer = &H100
    Private Const VK_C As Integer = &H43 ' A 키의 가상 키 코드

    ' Windows API 함수 선언
    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function SetWindowsHookEx(ByVal idHook As Integer, ByVal lpfn As LowLevelKeyboardProc, ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Function GetForegroundWindow() As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As UInteger
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function

    Private Delegate Function LowLevelKeyboardProc(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr

    Private hookID As IntPtr = IntPtr.Zero
    Private keyboardProc As LowLevelKeyboardProc = Nothing ' 대리자를 저장할 변수 추가

    Public Sub StartGlobalKeyboardHook()
        ' 키보드 후킹 시작
        keyboardProc = AddressOf HookCallback ' 대리자를 변수에 할당하여 가비지 수집을 방지합니다
        hookID = SetHook(keyboardProc)
    End Sub

    Public Sub StopGlobalKeyboardHook()
        ' 키보드 후킹 종료
        UnhookWindowsHookEx(hookID)
        keyboardProc = Nothing ' 대리자를 해제합니다
    End Sub

    Private Function SetHook(ByVal proc As LowLevelKeyboardProc) As IntPtr
        Using curProcess As Process = Process.GetCurrentProcess()
            Using curModule As ProcessModule = curProcess.MainModule
                Return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0)
            End Using
        End Using
    End Function

    Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso wParam = CType(WM_KEYDOWN, IntPtr) Then
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)
            If vkCode = VK_C AndAlso My.Computer.Keyboard.CtrlKeyDown Then
                ' 현재 프로세스가 활성 창인지 확인
                Dim foregroundHwnd As IntPtr = GetForegroundWindow()
                Dim foregroundProcessId As Integer = 0
                GetWindowThreadProcessId(foregroundHwnd, foregroundProcessId)

                If foregroundProcessId = Process.GetCurrentProcess().Id Then
                    ' 활성 창이 현재 프로세스일 때만 작업 수행
                    'MessageBox.Show("Ctrl + C가 눌렸습니다.")

                    Form1.ListView1.Columns.Clear()
                    Form1.ListView1.Items.Clear()
                    Form1.Textbox1.Enabled = True
                    Form1.TextBox2.Enabled = False
                End If
            End If
        End If
        Return CallNextHookEx(IntPtr.Zero, nCode, wParam, lParam)
    End Function
End Module
