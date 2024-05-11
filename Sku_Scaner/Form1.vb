Imports ExcelDataReader
Imports OfficeOpenXml ' EPPlus 라이브러리를 사용하기 위한 네임스페이스
Imports System.IO ' 파일 처리를 위한 네임스페이스


Public Class Form1

    Dim FilePath As String = "D:\sku_Scaner\data.xlsx" ' 엑셀 파일 경로
    Dim itemCount As Integer = 0 ' 상품 수량 카운트 변수 전역 변수로 변경
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        ToolStripStatusLabel1.Text = FilePath
        Textbox1.Enabled = True
        TextBox2.Enabled = False
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
            e.Item.BackColor = Color.Yellow
        Else
            e.Item.BackColor = Color.White
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
                foundItem.BackColor = Color.Yellow
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

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        ' TextBox2의 텍스트에서 공백을 제거
        Dim trimmedText As String = TextBox2.Text.Trim()

        ' 텍스트 길이가 13이 아니면 함수를 즉시 종료
        If Not trimmedText.Length = 13 Then
            TextBox2.Text = ""
            Exit Sub
        End If

        ' Enter 키를 눌렀을 때만 실행
        If e.KeyCode = Keys.Enter Then

            Listview_check()
            TextBox2.Text = ""
        End If
    End Sub

    Private Sub Textbox1_KeyDown(sender As Object, e As KeyEventArgs) Handles Textbox1.KeyDown
        ' Textbox1의 텍스트에서 공백을 제거
        Dim trimmedText As String = Textbox1.Text.Trim()
        ' 텍스트 길이가 12이 아니면 함수를 즉시 종료
        If Not trimmedText.Length = 12 Then
            Exit Sub
        End If

        ' Enter 키를 누르고, Textbox1의 텍스트 길이가 공백 제거 후 12이면 scan_start 메서드 실행
        If e.KeyCode = Keys.Enter Then  '송장번호만 입력받기
            scan_start()

            Textbox1.Enabled = False
            TextBox2.Enabled = True
            TextBox2.Focus()
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

End Class
