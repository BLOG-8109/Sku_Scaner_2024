﻿Imports ExcelDataReader
Imports OfficeOpenXml ' EPPlus 라이브러리를 사용하기 위한 네임스페이스
Imports System.IO ' 파일 처리를 위한 네임스페이스
Imports System.Media
Imports System.Numerics
Imports System.Resources
Imports System.Reflection
Imports System.Collections
Imports Microsoft.VisualBasic



Public Class Form1

    Dim FilePath As String = Application.StartupPath & "\data.xlsx"
    Dim itemCount As Integer = 0 ' 상품 수량 카운트 변수 전역 변수로 변경

    Private currentIndex As Integer = 0
    Private mp3Files As List(Of String)
    Private player As SoundPlayer
    Private resourceManager As ResourceManager

    Dim Channel As Integer = 0

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        'ToolStripStatusLabel1.Text = FilePath
        Textbox1.Enabled = True
        TextBox2.Enabled = False

        ' 여기에서 mp3 파일 목록을 초기화하거나, 외부에서 이 함수를 호출하여 설정
        mp3Files = New List(Of String) From {
        Application.StartupPath & "\start.wav",
        Application.StartupPath & "\Beep.wav",
        Application.StartupPath & "\end.wav"
}

        StartGlobalKeyboardHook()
    End Sub
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        StopGlobalKeyboardHook()
    End Sub

    Private Sub AddResourceDataUsingInputBox(resxFilePath As String)
        ' 사용자로부터 키와 값을 입력받음
        Dim key As String = InputBox("연동코드 입력", "연동코드 입력")
        Dim value As String = InputBox("바코드 입력", "바코드 입력")

        ' 입력받은 키와 값이 유효한지 확인
        If String.IsNullOrEmpty(key) OrElse String.IsNullOrEmpty(value) Then
            MessageBox.Show("Key or value cannot be empty.")
            Return
        End If

        ' ResXResourceWriter 객체 생성
        Using writer As New ResXResourceWriter(resxFilePath)
            ' 기존 리소스 데이터 불러오기
            If System.IO.File.Exists(resxFilePath) Then
                Using reader As New ResXResourceReader(resxFilePath)
                    reader.UseResXDataNodes = True
                    Dim node As DictionaryEntry
                    For Each node In reader
                        Dim resxNode As ResXDataNode = CType(node.Value, ResXDataNode)
                        writer.AddResource(resxNode)  ' 기존 데이터를 새 writer에 추가
                    Next
                End Using
            End If

            ' 새로운 리소스 데이터 추가
            writer.AddResource(key, value)

            ' 리소스 파일에 변경사항 저장
            writer.Generate()
        End Using
        MessageBox.Show("추가 완료")
    End Sub

    Private Sub play_wav(ByVal idx As Integer)
        player = New SoundPlayer(mp3Files(idx))
        player.Load()
        player.Play()
    End Sub
    Private Sub OpenFileMenuItem_Click(sender As Object, e As EventArgs) _
    Handles 열기EzAdminToolStripMenuItem.Click, 열기ShopeeToolStripMenuItem.Click, 열기Qoo10ToolStripMenuItem.Click

        OpenFile()
        ' sender를 ToolStripMenuItem으로 캐스팅하여 메뉴 아이템 참조
        Dim menuItem As ToolStripMenuItem = CType(sender, ToolStripMenuItem)
        ' 클릭된 메뉴 아이템에 따라 Channel 설정
        Select Case menuItem.Name
            Case "열기EzAdminToolStripMenuItem"
                Channel = 0
            Case "열기ShopeeToolStripMenuItem"
                Channel = 1
            Case "열기Qoo10ToolStripMenuItem"
                Channel = 2
        End Select
    End Sub
    Private Sub SaveTextToDateFile()
        Dim currentDate As String = DateTime.Now.ToString("yyyyMMdd") ' "yyyyMMdd" 형식으로 날짜를 문자열로 변환
        Dim filePath As String = Application.StartupPath & "\" & currentDate & ".txt" ' 파일 이름에 날짜 포함
        Using writer As StreamWriter = New StreamWriter(filePath, True) ' 파일에 내용을 추가
            writer.WriteLine(Textbox1.Text) ' TextBox1의 텍스트를 파일에 쓰기
        End Using
    End Sub
    Private Function CheckForDuplicatesToday() As Boolean
        Dim currentDate As String = DateTime.Now.ToString("yyyyMMdd")
        Dim filePath As String = Application.StartupPath & "\" & currentDate & ".txt"
        If File.Exists(filePath) Then
            Using reader As StreamReader = New StreamReader(filePath)
                Dim line As String
                Do
                    line = reader.ReadLine()
                    If line IsNot Nothing AndAlso line.Equals(Textbox1.Text, StringComparison.OrdinalIgnoreCase) Then
                        Return True ' 중복 발견 시 True 반환
                    End If
                Loop Until line Is Nothing
            End Using
        End If
        Return False ' 중복 없음
    End Function



    Private Sub scan_start()
        ' 엑셀 파일 경로
        Dim fileInfo As New FileInfo(FilePath)
        Using package As New ExcelPackage(fileInfo)
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)

            ' ListView 설정 초기화
            ListView1.Columns.Clear()
            ListView1.Items.Clear()

            ' 직접 지정한 헤더를 배열로 정의합니다.
            Dim customHeaders() As String = {"송장번호", "바코드", "연동코드", "상품명", "수량", "검수수량"}

            ' ListView의 열을 설정합니다.
            For Each header As String In customHeaders
                ListView1.Columns.Add(header)
            Next

            ' 사용자가 입력한 조건을 텍스트 상자에서 읽어옴
            Dim condition As String = Textbox1.Text.Trim()

            ' Dictionary를 사용하여 바코드를 키로, 해당 바코드의 총 수량과 기타 정보를 값으로 저장
            Dim barcodeInfo As New Dictionary(Of String, (Integer, String, String))

            For row As Integer = 2 To worksheet.Dimension.End.Row ' 1행은 헤더이므로 2행부터 시작
                ' 사용자가 입력한 조건과 현재 행의 첫 번째 열 값이 일치하는지 확인
                If worksheet.Cells(row, 2).Text = condition Then
                    Dim barcode As String = Trim(worksheet.Cells(row, 3).Text) ' 바코드 가져오기
                    Dim productName As String = worksheet.Cells(row, 4).Text ' 상품명 가져오기
                    Dim quantity As Integer = Integer.Parse(worksheet.Cells(row, 5).Text) ' 수량 가져오기
                    Dim linkageCode As String = worksheet.Cells(row, 6).Text ' 연동코드 가져오기

                    ' Dictionary에 해당 바코드가 이미 있는지 확인하고, 없으면 추가하고 있으면 수량을 더함
                    If Not barcodeInfo.ContainsKey(barcode) Then
                        barcodeInfo.Add(barcode, (quantity, productName, linkageCode))
                    Else
                        Dim currentInfo = barcodeInfo(barcode)
                        barcodeInfo(barcode) = (currentInfo.Item1 + quantity, currentInfo.Item2, currentInfo.Item3)
                    End If
                End If
            Next

            ' ListView에 합쳐진 데이터를 표시
            For Each kvp As KeyValuePair(Of String, (Integer, String, String)) In barcodeInfo
                Dim newRow As New ListViewItem(condition) ' 조건 추가
                newRow.SubItems.Add(kvp.Key) ' 바코드 추가
                newRow.SubItems.Add(kvp.Value.Item3) ' 연동코드 추가
                newRow.SubItems.Add(kvp.Value.Item2) ' 상품명 추가
                newRow.SubItems.Add(kvp.Value.Item1.ToString()) ' 총 수량 추가
                newRow.SubItems.Add(0) ' 연동코드 추가
                ListView1.Items.Add(newRow)
            Next


        End Using

        ' 모든 컬럼 너비 자동 조절
        For Each column As ColumnHeader In ListView1.Columns
            column.Width = -2
        Next


    End Sub
    Private Sub shopee_start()
        ' 엑셀 파일 경로
        Dim fileInfo As New FileInfo(FilePath)
        Using package As New ExcelPackage(fileInfo)
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)

            ' ListView 설정 초기화
            ListView1.Columns.Clear()
            ListView1.Items.Clear()

            ' 직접 지정한 헤더를 배열로 정의합니다.
            Dim customHeaders() As String = {"송장번호", "바코드", "연동코드", "상품명", "수량", "검수수량"}

            ' ListView의 열을 설정합니다.
            For Each header As String In customHeaders
                ListView1.Columns.Add(header)
            Next

            ' 사용자가 입력한 조건을 텍스트 상자에서 읽어옴
            Dim condition As String = Textbox1.Text.Trim()

            ' Dictionary를 사용하여 바코드를 키로, 해당 바코드의 총 수량과 기타 정보를 값으로 저장
            Dim barcodeInfo As New Dictionary(Of String, (Integer, String, String))
            ' 리소스 매니저 초기화
            resourceManager = New ResourceManager("Sku_Scaner.barcode_data", GetType(Form1).Assembly)


            For row As Integer = 2 To worksheet.Dimension.End.Row ' 1행은 헤더이므로 2행부터 시작
                ' 사용자가 입력한 조건과 현재 행의 첫 번째 열 값이 일치하는지 확인
                If worksheet.Cells(row, 2).Text = condition Then
                    Dim barcode As String = resourceManager.GetString(worksheet.Cells(row, 6).Text)
                    Dim productName As String = worksheet.Cells(row, 4).Text ' 상품명 가져오기
                    Dim quantity As Integer = Integer.Parse(worksheet.Cells(row, 5).Text) ' 수량 가져오기
                    Dim linkageCode As String = worksheet.Cells(row, 6).Text ' 연동코드 가져오기

                    ' Dictionary에 해당 바코드가 이미 있는지 확인하고, 없으면 추가하고 있으면 수량을 더함
                    If Not barcodeInfo.ContainsKey(barcode) Then
                        barcodeInfo.Add(barcode, (quantity, productName, linkageCode))
                    Else
                        Dim currentInfo = barcodeInfo(barcode)
                        barcodeInfo(barcode) = (currentInfo.Item1 + quantity, currentInfo.Item2, currentInfo.Item3)
                    End If
                End If
            Next

            ' ListView에 합쳐진 데이터를 표시
            For Each kvp As KeyValuePair(Of String, (Integer, String, String)) In barcodeInfo
                Dim newRow As New ListViewItem(condition) ' 조건 추가
                newRow.SubItems.Add(kvp.Key) ' 바코드 추가
                newRow.SubItems.Add(kvp.Value.Item3) ' 연동코드 추가
                newRow.SubItems.Add(kvp.Value.Item2) ' 상품명 추가
                newRow.SubItems.Add(kvp.Value.Item1.ToString()) ' 총 수량 추가
                newRow.SubItems.Add(0) ' 연동코드 추가
                ListView1.Items.Add(newRow)
            Next


        End Using

        ' 모든 컬럼 너비 자동 조절
        For Each column As ColumnHeader In ListView1.Columns
            column.Width = -2
        Next


    End Sub
    Private Sub OpenFile()
        ' 파일을 열기 위한 OpenFileDialog 생성
        Dim openFileDialog As New OpenFileDialog()

        ' 사용자가 파일을 선택하고 확인을 누르면
        If openFileDialog.ShowDialog() = DialogResult.OK Then
            ' 선택된 파일의 경로
            FilePath = openFileDialog.FileName

            ' 파일의 확장자를 확인하여 처리
            If Path.GetExtension(FilePath).ToLower() = ".xls" Then
                ' XLS 파일을 XLSX 파일로 변환하여 처리
                ConvertXlsToXlsx(FilePath, Application.StartupPath & "\data.xlsx")
                FilePath = Application.StartupPath & "\data.xlsx"
                CountUniqueValuesInColumnB()
                'OpenAndProcessFile(convertedFilePath)
            Else
                CountUniqueValuesInColumnB()
                ' 그 외의 경우에는 그대로 파일을 처리
                'OpenAndProcessFile(selectedFilePath)
            End If
        End If
        ' ToolStripStatusLabel1.Text = FilePath

    End Sub
    Private Sub CountUniqueValuesInColumnB()
        ' 파일 경로 설정
        'Dim filePath As String = "D:\sku_Scaner\data.xlsx"
        Dim fileInfo As New FileInfo(filePath)

        ' ExcelPackage 객체를 사용하여 파일 열기
        Using package As New ExcelPackage(fileInfo)
            ' 첫 번째 워크시트 가져오기
            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets(0)

            ' B열 데이터를 저장할 HashSet 생성 (HashSet은 중복을 자동으로 제거)
            Dim uniqueValues As New HashSet(Of String)

            ' 2행부터 시작하여 마지막 행까지 반복 (1행은 헤더일 경우)
            For row As Integer = 2 To worksheet.Dimension.End.Row
                ' 현재 행의 B열 데이터 읽기 (B열은 2번째 열)
                Dim value As String = worksheet.Cells(row, 2).Text

                ' HashSet에 값 추가 (중복 값은 자동으로 무시됨)
                uniqueValues.Add(value)
            Next

            ' 고유 값의 개수 출력
            ToolStripStatusLabel1.Text = "총 주문 건 수 : " & uniqueValues.Count - 1
            ToolStripProgressBar1.Maximum = uniqueValues.Count - 1
            'Console.WriteLine("Total unique values in column B: " & uniqueValues.Count)
        End Using
    End Sub


    Private Sub ConvertXlsToXlsx(ByVal xlsPath As String, ByVal xlsxPath As String)
        ' 파일 스트림을 사용하여 .XLS 파일을 열기
        Using stream As FileStream = File.Open(xlsPath, FileMode.Open, FileAccess.Read)
            ' ExcelDataReader를 사용하여 데이터 읽기
            Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                ' DataSet으로 데이터 가져오기
                Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                    .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                        .UseHeaderRow = True
                    }
                })

                ' EPPlus를 사용하여 새 Excel 파일 생성
                Using package As New ExcelPackage()
                    For Each table As DataTable In result.Tables
                        ' 각 DataTable을 새 시트로 추가
                        Dim worksheet = package.Workbook.Worksheets.Add(table.TableName)
                        worksheet.Cells("A1").LoadFromDataTable(table, True)
                    Next

                    ' .XLSX 형식으로 저장
                    File.WriteAllBytes(xlsxPath, package.GetAsByteArray())
                End Using
            End Using
        End Using
    End Sub


    Private Sub ListView1_ItemChecked(sender As Object, e As ItemCheckedEventArgs) Handles ListView1.ItemChecked
        If e.Item.Checked Then
            e.Item.BackColor = System.Drawing.Color.Yellow
        Else
            e.Item.BackColor = System.Drawing.Color.White
        End If
        ' 모든 아이템이 체크되었는지 다시 검사
        CheckItemsAndPerformActions()
    End Sub

    Private Sub Listview_check()
        ' TextBox2에서 입력된 바코드 값
        Dim barcode As String = Trim(TextBox2.Text)

        Dim foundItem As ListViewItem = Nothing
        ' ListView의 모든 항목을 검사
        For Each item As ListViewItem In ListView1.Items
            ' 입력된 바코드와 ListView의 바코드가 일치하는지 검사
            If item.SubItems(1).Text = barcode Then
                ' 해당 항목이 이미 체크되었는지 검사
                If item.Checked Then
                    play_wav(1) ' beep wav
                    MessageBox.Show("이미 검수 완료된 상품입니다.")
                    Exit Sub
                Else
                    foundItem = item
                    Exit For
                End If
            End If
        Next

        ' 일치하는 항목을 찾았을 경우
        If foundItem IsNot Nothing Then
            ' 검수 수량 업데이트
            Dim count As Integer = Integer.Parse(foundItem.SubItems(5).Text)
            count += 1
            foundItem.SubItems(5).Text = count.ToString()

            ' 검수 수량이 주문 수량과 동일하다면 항목 체크
            If count >= Integer.Parse(foundItem.SubItems(4).Text) Then
                foundItem.Checked = True
                foundItem.BackColor = System.Drawing.Color.Yellow
                itemCount = 0 ' itemCount 초기화



            End If
        Else
            play_wav(1) ' beep wav
            MessageBox.Show("일치하는 항목이 없습니다.")
        End If
    End Sub


    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        Dim trimmedText As String = TextBox2.Text.Trim() '공백 제거

        If e.KeyChar = Convert.ToChar(Keys.Enter) Then
            If trimmedText.Length = 13 Then '13자리 바코드 일때만 작동
                Listview_check()
                TextBox2.Text = ""
                e.Handled = True

            Else
                play_wav(1)
                TextBox2.Text = ""
                e.Handled = True
            End If
        End If
    End Sub

    Private Sub Textbox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Textbox1.KeyPress

        If CheckForDuplicatesToday() And e.KeyChar = Convert.ToChar(Keys.Enter) Then
            Dim result As DialogResult = MessageBox.Show("이미 검수 완료된 송장입니다. 다시 검수 하시겠습니까?", "중복 검사", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.No Then
                Exit Sub
            End If
        End If

        Dim trimmedText As String = Textbox1.Text.Trim()

        ' Enter 키를 누르고, Textbox1의 텍스트 길이가 공백 제거 후 12 또는 15자리일 때 처리
        If e.KeyChar = Convert.ToChar(Keys.Enter) Then
            ' 12자리 또는 15자리 송장번호만 입력받기
            If trimmedText.Length = 12 Or trimmedText.Length = 15 Then
                play_wav(0) ' 시작 wav
                ' Channel 값에 따라 다른 시작 메서드 호출
                Select Case Channel
                    Case 0 ' EzAdmin
                        scan_start()
                    Case 1 ' Shopee
                        shopee_start()
                    Case 2 ' Qoo10
                        'qoo10_start() ' 가정: Qoo10에 대한 처리 메서드가 존재한다고 가정
                End Select
                Textbox1.Enabled = False
                TextBox2.Enabled = True
                TextBox2.Focus()
            Else
                play_wav(1) ' beep wav
                Textbox1.Text = vbNullString
            End If
            e.Handled = True
        End If
    End Sub


    'Private Function AllItemsChecked() As Boolean
    '    For Each item As ListViewItem In ListView1.Items
    '        If Not item.Checked Then
    '            Return False ' 하나라도 체크되어 있지 않으면 False 반환
    '        End If
    '    Next
    '    Return True ' 모두 체크되어 있으면 True 반환
    'End Function

    Private Sub CheckItemsAndPerformActions()
        ' 모든 아이템 체크 여부 검사 및 결과에 따라 액션 수행
        If ListView1.Items.Cast(Of ListViewItem).All(Function(item) item.Checked) Then
            PerformAllCheckedActions()
        End If
    End Sub

    Private Sub PerformAllCheckedActions()
        ' 모든 아이템이 체크된 경우 실행할 작업들
        SaveTextToDateFile() ' TXT 파일 저장
        UpdateStatus()        ' 상태 업데이트
        ResetForm()           ' 폼 초기화
    End Sub

    Private Sub UpdateStatus()
        On Error Resume Next
        ' 상태바 업데이트
        ToolStripProgressBar1.Value += 1
        ToolStripStatusLabel1.Text = $"{ToolStripProgressBar1.Value}/{ToolStripProgressBar1.Maximum}"
        play_wav(2) ' 종료 신호음 재생
    End Sub

    Private Sub ResetForm()
        ' 폼 컨트롤 초기화
        With ListView1
            .Columns.Clear()
            .Items.Clear()
        End With
        With Textbox1
            .Enabled = True
            .Text = vbNullString
        End With
        With TextBox2
            .Enabled = False
            .Text = vbNullString
        End With
    End Sub

    Private Sub ListView1_MouseWheel(sender As Object, e As MouseEventArgs) Handles ListView1.MouseWheel
        Dim delta As Integer = e.Delta

        ' 폰트 크기 증가
        If delta > 0 Then
            For Each item As ListViewItem In ListView1.Items
                item.Font = New Font(item.Font.FontFamily, item.Font.Size + 1)
            Next
        Else ' 폰트 크기 감소
            For Each item As ListViewItem In ListView1.Items
                ' 폰트 크기가 1보다 작아지지 않도록 함
                If item.Font.Size > 1 Then
                    item.Font = New Font(item.Font.FontFamily, item.Font.Size - 1)
                End If
            Next
        End If

        ' 항목 높이 및 너비 조정
        AdjustListViewItemSize()
    End Sub

    Private Sub AdjustListViewItemSize()
        ' 컬럼 폭에 따라 ListView의 너비 조정
        For Each column As ColumnHeader In ListView1.Columns
            column.Width = -2 ' AutoResize
        Next

        ' ListView의 각 항목 높이 조정
        For Each item As ListViewItem In ListView1.Items
            item.Selected = True ' 선택된 항목의 높이가 변경되도록 함
            item.Selected = False
        Next
    End Sub

    Private Sub 바코드추가ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 바코드추가ToolStripMenuItem.Click
        Dim resxPath As String = Application.StartupPath & "\barcode_data.resx"
        'MsgBox(resxPath)
        AddResourceDataUsingInputBox(resxPath)
    End Sub

End Class
