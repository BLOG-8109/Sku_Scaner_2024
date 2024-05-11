﻿Imports ExcelDataReader
Imports OfficeOpenXml ' EPPlus 라이브러리를 사용하기 위한 네임스페이스
Imports System.IO ' 파일 처리를 위한 네임스페이스
Imports System.Media
Imports System.Numerics


Public Class Form1

    Dim FilePath As String = "D:\sku_Scaner\data.xlsx" ' 엑셀 파일 경로
    Dim itemCount As Integer = 0 ' 상품 수량 카운트 변수 전역 변수로 변경

    Private currentIndex As Integer = 0
    Private mp3Files As List(Of String)
    Private player As SoundPlayer
    'Private mediaPlayer As New WMPLib.WindowsMediaPlayer()
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        ToolStripStatusLabel1.Text = FilePath
        Textbox1.Enabled = True
        TextBox2.Enabled = False

        ' 여기에서 mp3 파일 목록을 초기화하거나, 외부에서 이 함수를 호출하여 설정
        mp3Files = New List(Of String) From {
            "D:\sku_Scaner\Sku_Scaner\fail.wav",
            "D:\sku_Scaner\Sku_Scaner\fail.mp3",
            "D:\sku_Scaner\Sku_Scaner\fail.mp3"}
        StartGlobalKeyboardHook()
    End Sub
    Private Sub Form1_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        StopGlobalKeyboardHook()
    End Sub
    Private Sub play_wav(ByVal idx As Integer)
        player = New SoundPlayer(mp3Files(idx))
        player.Load()
        player.Play()
    End Sub
    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
    '    ToolStripStatusLabel1.Text = DateTime.Now.ToString("HH:mm:ss")
    'End Sub

    Private Sub 열기EzAdminToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 열기EzAdminToolStripMenuItem.Click
        OpenFile()
    End Sub

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
            Dim condition As String = Textbox1.Text

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
                ConvertXlsToXlsx(FilePath, "D:\sku_Scaner\data.xlsx")
                FilePath = "D:\sku_Scaner\data.xlsx"
                'OpenAndProcessFile(convertedFilePath)
            Else
                ' 그 외의 경우에는 그대로 파일을 처리
                'OpenAndProcessFile(selectedFilePath)
            End If
        End If
        ToolStripStatusLabel1.Text = FilePath
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
    End Sub

    Private Sub Listview_check()
        ' 모든 아이템이 체크되어 있지 않으면 종료
        If AllItemsChecked() Then
            'Return
            ' ListView 설정 초기화
            ListView1.Columns.Clear()
            ListView1.Items.Clear()
        End If

        Dim barcode As String = Trim(TextBox2.Text)

        Dim foundItem As ListViewItem = Nothing
        For Each item As ListViewItem In ListView1.Items
            If item.SubItems(1).Text = barcode Then
                If item.Checked Then

                    MessageBox.Show("이미 검수 완료된 상품입니다.")

                    Exit Sub
                Else
                    foundItem = item
                    Exit For
                End If
            End If
        Next

        If foundItem IsNot Nothing Then
            itemCount += 1
            foundItem.SubItems(5).Text = itemCount
            If itemCount >= Integer.Parse(foundItem.SubItems(4).Text) Then
                foundItem.Checked = True
                foundItem.BackColor = System.Drawing.Color.Yellow
                itemCount = 0

                ' 모든 아이템이 체크되어 있지 않으면 종료
                If AllItemsChecked() Then
                    'Return
                    ' ListView 설정 초기화
                    ListView1.Columns.Clear()
                    ListView1.Items.Clear()
                    Textbox1.Enabled = True
                    TextBox2.Enabled = False
                End If

            End If
        Else
            MessageBox.Show("일치하는 항목이 없습니다.")
            Exit Sub
        End If

    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        'e.Handled = True
        Dim trimmedText As String = TextBox2.Text.Trim()
        If e.KeyChar = Convert.ToChar(Keys.Enter) AndAlso trimmedText.Length = 13 Then

            Listview_check()
            TextBox2.Text = ""
            e.Handled = True
        End If
    End Sub

    Private Sub Textbox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Textbox1.KeyPress
        Dim trimmedText As String = Textbox1.Text.Trim()
        'e.Handled = True
        ' Enter 키를 누르고, Textbox1의 텍스트 길이가 공백 제거 후 12이면 scan_start 메서드 실행
        If e.KeyChar = Convert.ToChar(Keys.Enter) Then  '송장번호만 입력받기
            If trimmedText.Length = 12 Then
                scan_start()

                Textbox1.Enabled = False
                TextBox2.Enabled = True
                TextBox2.Focus()
                e.Handled = True
            End If
        End If
    End Sub

    Private Function AllItemsChecked() As Boolean
        For Each item As ListViewItem In ListView1.Items
            If Not item.Checked Then
                Return False ' 하나라도 체크되어 있지 않으면 False 반환
            End If
        Next
        Return True ' 모두 체크되어 있으면 True 반환
    End Function
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

End Class
