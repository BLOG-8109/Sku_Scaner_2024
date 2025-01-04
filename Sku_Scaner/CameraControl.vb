Imports System.Diagnostics
Imports System.IO

Module CameraControl
    Sub Main()
        Try
            ' 오늘 날짜를 "yyyyMMdd" 형식으로 가져오기
            Dim todayDate As String = DateTime.Now.ToString("yyyyMMdd")

            ' 바탕화면 경로를 동적으로 얻기
            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

            ' 바탕화면에 오늘 날짜 폴더 만들기
            Dim folderPath As String = Path.Combine(desktopPath, todayDate)
            CreateDirectoryIfNotExists(folderPath)

            ' 카메라 촬영 버튼 클릭
            ExecuteAdbCommand("shell input keyevent 27", "사진 촬영 완료")

            ' 파일 경로 목록 검색
            Dim findCommand As String = $"shell find /sdcard/DCIM/Camera/ -type f -name ""{todayDate}_*.jpg"""
            Dim output As String = ExecuteAdbCommand(findCommand, "사진 목록 검색 완료", redirectOutput:=True)

            ' 사진 파일이 없으면 종료
            If String.IsNullOrEmpty(output) Then
                Console.WriteLine("사진을 찾을 수 없습니다.")
                MsgBox("?")
                Return
            End If

            ' 파일 목록 분리
            Dim filePaths = output.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            ' 각 파일을 복사 및 삭제
            For Each filePath In filePaths
                ' 파일 복사
                ExecuteAdbCommand($"pull {filePath} {folderPath}", $"파일 복사 완료: {filePath}")

                ' 원본 파일 삭제
                ExecuteAdbCommand($"shell rm {filePath}", $"원본 파일 삭제 완료: {filePath}")
            Next

            ' 모든 작업 완료 후 알림
            Console.WriteLine("모든 파일이 오늘 날짜 폴더로 옮겨지고 원본이 삭제되었습니다!")
        Catch ex As Exception
            ' 예외 발생 시 에러 메시지 출력
            Console.WriteLine($"에러 발생: {ex.Message}")
        End Try
    End Sub

    ' 프로세스를 실행하는 함수
    Private Function ExecuteAdbCommand(arguments As String, successMessage As String, Optional redirectOutput As Boolean = False) As String
        Dim process As New Process()
        process.StartInfo.FileName = "C:\adb\adb.exe"
        process.StartInfo.Arguments = arguments
        process.StartInfo.UseShellExecute = False
        process.StartInfo.CreateNoWindow = True
        process.StartInfo.RedirectStandardOutput = redirectOutput

        process.Start()
        Dim output As String = If(redirectOutput, process.StandardOutput.ReadToEnd().Trim(), String.Empty)
        process.WaitForExit()

        Console.WriteLine(successMessage)
        Return output
    End Function

    ' 디렉터리가 없으면 생성하는 함수
    Private Sub CreateDirectoryIfNotExists(folderPath As String)
        If Not Directory.Exists(folderPath) Then
            Directory.CreateDirectory(folderPath)
        End If
    End Sub
End Module
