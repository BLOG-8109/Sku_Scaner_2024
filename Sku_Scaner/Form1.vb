Imports OfficeOpenXml
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1

    Dim filePath As String = "D:\sku_Scaner\data2024.xlsx" ' 엑셀 파일 경로

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel1.Text = DateTime.Now.ToString("HH:mm:ss")
    End Sub

    Private Sub 열기EzAdminToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles 열기EzAdminToolStripMenuItem.Click
        OpenFile()

    End Sub

    Private Sub scan_start()
        ' 엑셀 파일 경로
        Dim fileInfo As New FileInfo(filePath)

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
                    Dim barcode As String = worksheet.Cells(row, 3).Text ' 바코드 가져오기
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
            Dim selectedFilePath As String = openFileDialog.FileName

            ' 파일의 확장자를 확인하여 처리
            If Path.GetExtension(selectedFilePath).ToLower() = ".xls" Then
                ' XLS 파일을 XLSX 파일로 변환하여 처리
                Dim convertedFilePath As String = ConvertXlsToXlsx(selectedFilePath)
                'OpenAndProcessFile(convertedFilePath)
            Else
                ' 그 외의 경우에는 그대로 파일을 처리
                'OpenAndProcessFile(selectedFilePath)
            End If
        End If
    End Sub
    Private Function ConvertXlsToXlsx(filePath As String) As String
        ' 변환된 파일의 경로
        Dim convertedFilePath As String = Path.ChangeExtension(filePath, ".xlsx")

        ' XLS 파일을 XLSX 파일로 변환하는 코드
        ' 여기에 변환하는 코드를 작성해야 합니다.

        ' 변환된 파일의 경로 반환
        MsgBox(convertedFilePath)
        Return convertedFilePath
    End Function

    Private Sub Textbox1_KeyUp(sender As Object, e As KeyEventArgs) Handles Textbox1.KeyUp
        If e.KeyCode = 13 Then
            scan_start()
        End If
    End Sub
End Class
