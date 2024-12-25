Imports System.Diagnostics
Imports System.IO

Module CameraControl
    Sub Main()
        Try
            ' 오늘 날짜를 "yyyyMMdd" 형식으로 가져오기
            ' 예: 20241225
            Dim todayDate As String = DateTime.Now.ToString("yyyyMMdd")

            ' 바탕화면 경로를 동적으로 얻기
            Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)

            ' 바탕화면에 오늘 날짜 폴더 만들기
            Dim folderPath As String = Path.Combine(desktopPath, todayDate)
            If Not Directory.Exists(folderPath) Then
                Directory.CreateDirectory(folderPath) ' 폴더가 없으면 생성
            End If

            ' 카메라 촬영 버튼 클릭
            Dim captureProcess As New Process()
            captureProcess.StartInfo.FileName = "C:\adb\adb.exe" ' adb 실행 파일 경로
            captureProcess.StartInfo.Arguments = "shell input keyevent 27" ' 촬영 버튼 클릭
            captureProcess.StartInfo.UseShellExecute = False ' 새로운 쉘을 사용하지 않도록 설정
            captureProcess.StartInfo.CreateNoWindow = True ' 명령 프롬프트 창을 숨기기 위해 설정
            captureProcess.Start() ' 프로세스 시작
            captureProcess.WaitForExit() ' 프로세스가 종료될 때까지 기다림

            ' 파일 경로 목록 검색
            Dim findProcess As New Process()
            findProcess.StartInfo.FileName = "C:\adb\adb.exe"
            findProcess.StartInfo.Arguments = $"shell find /sdcard/DCIM/Camera/ -type f -name ""{todayDate}_*.jpg"""
            findProcess.StartInfo.RedirectStandardOutput = True ' 명령 출력값을 가져올 수 있도록 설정
            findProcess.StartInfo.UseShellExecute = False
            findProcess.StartInfo.CreateNoWindow = True ' 명령 프롬프트 창을 숨기기 위해 설정
            findProcess.Start() ' 프로세스 시작
            Dim output As String = findProcess.StandardOutput.ReadToEnd().Trim() ' 결과값을 읽어옴
            findProcess.WaitForExit() ' 프로세스 종료 대기

            ' 사진 파일이 없으면 종료
            If String.IsNullOrEmpty(output) Then
                Console.WriteLine("사진을 찾을 수 없습니다.") ' 사진을 찾지 못한 경우 메시지 출력
                Return ' 프로그램 종료
            End If

            ' 파일 목록 분리 및 복사
            ' find 명령으로 검색된 파일 경로를 줄바꿈을 기준으로 분리하여 배열에 저장
            Dim filePaths = output.Split(New String() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)

            ' 각 파일을 오늘 날짜 폴더로 복사
            ' "adb pull" 명령을 사용하여 스마트폰에서 PC로 파일을 복사
            For Each filePath In filePaths
                Dim pullProcess As New Process()
                pullProcess.StartInfo.FileName = "C:\adb\adb.exe"
                pullProcess.StartInfo.Arguments = $"pull {filePath} {folderPath}" ' 오늘 날짜 폴더로 복사
                pullProcess.StartInfo.RedirectStandardOutput = True
                pullProcess.StartInfo.UseShellExecute = False
                pullProcess.StartInfo.CreateNoWindow = True ' 명령 프롬프트 창을 숨기기 위해 설정
                pullProcess.Start() ' 프로세스 시작
                pullProcess.WaitForExit() ' 프로세스 종료 대기
                Console.WriteLine($"파일 복사 완료: {filePath}") ' 복사 완료된 파일 경로 출력
            Next

            ' 모든 파일 복사가 완료된 후 알림
            Console.WriteLine("모든 파일이 오늘 날짜 폴더로 옮겨졌습니다!")
        Catch ex As Exception
            ' 예외 발생 시 에러 메시지 출력
            Console.WriteLine($"에러 발생: {ex.Message}")
        End Try
    End Sub
End Module
