
Option Explicit

' ====== Cấu hình ======
Private Const OPTIONS_RANGE As String = "A2:A50" ' Vùng chứa lựa chọn (mỗi ô một lựa chọn)
Private Const WHEEL_CHART_NAME As String = "WheelChart"
Private Const POINTER_NAME As String = "WheelPointer"
Private Const POINTER_ANGLE As Double = 90          ' Kim chỉ ở vị trí 12 giờ (90° theo Excel)
Private Const CHART_LEFT As Double = 200
Private Const CHART_TOP As Double = 50
Private Const CHART_WIDTH As Double = 400
Private Const CHART_HEIGHT As Double = 400

' ====== Tạo vòng quay từ danh sách ======
Public Sub BuildWheel()
    Dim ws As Worksheet
    Dim rng As Range, cell As Range
    Dim labels() As String
    Dim values() As Double
    Dim n As Long, i As Long
    
    Set ws = ActiveSheet
    On Error Resume Next
    Set rng = ws.Range(OPTIONS_RANGE).SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    
    If rng Is Nothing Then
        MsgBox "Không thấy dữ liệu lựa chọn trong vùng " & OPTIONS_RANGE, vbExclamation
        Exit Sub
    End If
    
    ' Đếm số lựa chọn (ô không rỗng)
    n = 0
    For Each cell In ws.Range(OPTIONS_RANGE)
        If Len(Trim(cell.Value)) > 0 Then n = n + 1
    Next cell
    If n < 2 Then
        MsgBox "Cần ít nhất 2 lựa chọn để tạo vòng quay.", vbInformation
        Exit Sub
    End If
    
    ' Nạp nhãn và giá trị đều nhau (1 cho mỗi lát)
    ReDim labels(1 To n)
    ReDim values(1 To n)
    i = 0
    For Each cell In ws.Range(OPTIONS_RANGE)
        If Len(Trim(cell.Value)) > 0 Then
            i = i + 1
            labels(i) = CStr(cell.Value)
            values(i) = 1
            If i = n Then Exit For
        End If
    Next cell
    
    ' Xóa chart/kim cũ (nếu có)
    DeleteIfExists ws, WHEEL_CHART_NAME
    DeleteIfExists ws, POINTER_NAME
    
    ' Tạo chart
    Dim co As ChartObject, ch As Chart
    Set co = ws.ChartObjects.Add(CHART_LEFT, CHART_TOP, CHART_WIDTH, CHART_HEIGHT)
    co.Name = WHEEL_CHART_NAME
    Set ch = co.Chart
    ch.ChartType = xlDoughnut   ' Có lỗ giữa sẽ đẹp hơn; có thể dùng xlPie
    
    ' Tạo series
    Dim ser As Series
    ch.SeriesCollection.NewSeries
    Set ser = ch.SeriesCollection(1)
    ser.Values = values
    ser.XValues = labels
    
    ' Ẩn legend, hiển thị nhãn bên trong lát
    ch.HasLegend = False
    ser.HasDataLabels = True
    ser.ApplyDataLabels xlDataLabelsShowLabel
    ser.DataLabels.Position = xlLabelPositionCenter
    
    ' Đổi màu lát cắt đa sắc
    Randomize
    Dim p As Point
    For i = 1 To n
        Set p = ser.Points(i)
        p.Format.Fill.ForeColor.RGB = RGB(Int(200 * Rnd) + 30, Int(200 * Rnd) + 30, Int(200 * Rnd) + 30)
        p.Format.Fill.Solid
    Next i
    
    ' Đặt góc lát đầu tiên
    ch.ChartGroups(1).FirstSliceAngle = 0
    
    ' Thêm kim chỉ ở đỉnh (12 giờ)
    AddPointer ws, co
    MsgBox "Đã tạo vòng quay. Dùng macro SpinWheel để quay!", vbInformation
End Sub

' ====== Quay vòng với hiệu ứng giảm tốc ======
Public Sub SpinWheel()
    Dim ws As Worksheet: Set ws = ActiveSheet
    Dim co As ChartObject, ch As Chart
    On Error Resume Next
    Set co = ws.ChartObjects(WHEEL_CHART_NAME)
    On Error GoTo 0
    If co Is Nothing Then
        MsgBox "Chưa có vòng quay. Hãy chạy BuildWheel trước.", vbExclamation
        Exit Sub
    End If
    Set ch = co.Chart
    
    ' Số lát
    Dim ser As Series
    Set ser = ch.SeriesCollection(1)
    Dim n As Long: n = ser.Points.Count
    If n < 2 Then
        MsgBox "Vòng quay cần ít nhất 2 lát.", vbExclamation
        Exit Sub
    End If
    
    ' Tổng góc quay ngẫu nhiên: 3–6 vòng + offset 0–360°
    Randomize
    Dim totalDegrees As Double
    totalDegrees = (360 * (3 + Int(3 * Rnd))) + (360 * Rnd)
    
    ' Hiệu ứng giảm tốc (ease-out cubic)
    Dim steps As Long: steps = 250
    Dim startAngle As Double: startAngle = ch.ChartGroups(1).FirstSliceAngle
    Dim i As Long, t As Double, eased As Double, newAngle As Double
    
    Application.ScreenUpdating = True
    For i = 1 To steps
        t = i / steps                    ' tiến độ 0→1
        eased = 1 - (1 - t) ^ 3          ' ease-out cubic
        newAngle = startAngle + eased * totalDegrees
        ch.ChartGroups(1).FirstSliceAngle = newAngle - 360 * Int(newAngle / 360) ' mod 360
        DoEvents
        DelayMs 8                        ' làm mượt chuyển động
    Next i
    
    ' Xác định lát trúng dưới kim chỉ
    Dim finalAngle As Double: finalAngle = ch.ChartGroups(1).FirstSliceAngle
    Dim sliceAngle As Double: sliceAngle = 360# / n
    Dim delta As Double
    
    ' Excel đo góc theo chiều kim đồng hồ từ hướng 3 giờ.
    ' Kim chỉ ở 12 giờ => góc kim = 90°
    delta = POINTER_ANGLE - finalAngle
    delta = delta - 360# * Int(delta / 360#)       ' mod 360, kết quả 0..360
    If delta < 0 Then delta = delta + 360#
    
    Dim index As Long
    index = Int(delta / sliceAngle) + 1
    If index < 1 Then index = 1
    If index > n Then index = n
    
    Dim winner As String
    winner = ser.XValues(index)
    
    ' Làm nổi bật lát trúng
    HighlightSlice ser, index
    
    MsgBox "KẾT QUẢ: " & winner, vbInformation, "Chiếc nón kỳ diệu"
End Sub

' ====== Tiện ích: thêm kim chỉ ======
Private Sub AddPointer(ws As Worksheet, co As ChartObject)
    Dim shp As Shape
    Dim cx As Double, cy As Double, w As Double, h As Double
    cx = co.Left + co.Width / 2
    cy = co.Top
    w = 30
    h = 40
    ' Tam giác cân chỉ xuống
    Set shp = ws.Shapes.AddShape(msoShapeIsoscelesTriangle, cx - w / 2, cy - h - 6, w, h)
    shp.Name = POINTER_NAME
    shp.Fill.ForeColor.RGB = RGB(220, 50, 50)
    shp.Line.Visible = msoFalse
    shp.Rotation = 0                 ' Hướng lên trên
End Sub

' ====== Tiện ích: làm nổi bật lát thắng ======
Private Sub HighlightSlice(ByVal ser As Series, ByVal idx As Long)
    Dim i As Long
    For i = 1 To ser.Points.Count
        ser.Points(i).Explosion = IIf(i = idx, 10, 0)  ' đẩy lát thắng ra một chút
        If i = idx Then
            ser.Points(i).Format.Line.Visible = msoTrue
            ser.Points(i).Format.Line.ForeColor.RGB = RGB(255, 255, 255)
            ser.Points(i).Format.Line.Weight = 2
        Else
            ser.Points(i).Format.Line.Visible = msoFalse
        End If
    Next i
End Sub

' ====== Tiện ích: xóa shape/chart nếu tồn tại ======
Private Sub DeleteIfExists(ws As Worksheet, ByVal name As String)
    On Error Resume Next
    ws.ChartObjects(name).Delete
    ws.Shapes(name).Delete
    On Error GoTo 0
End Sub

' ====== Tiện ích: delay mượt, không đóng băng ứng dụng ======
Private Sub DelayMs(ByVal ms As Long)
    Dim t0 As Single: t0 = Timer
    Do While (Timer - t0) * 1000 < ms
        DoEvents
    Loop
End Sub
