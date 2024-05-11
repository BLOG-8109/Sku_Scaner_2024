Imports OfficeOpenXml
Imports System.IO
Imports System.Windows.Forms.VisualStyles.VisualStyleElement

Public Class Form1

    Dim filePath As String = "D:\sku_Scaner\data2024.xlsx" ' 엑셀 파일 경로

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial

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

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        ToolStripStatusLabel1.Text = DateTime.Now.ToString("HH:mm:ss")

    End Sub
End Class
