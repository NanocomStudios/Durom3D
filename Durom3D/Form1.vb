Public Class Form1
    Dim Fheight As Integer
    Dim Fwidth As Integer

    Dim bitmap() As Byte

    Dim iscapture As Boolean = False
    Dim xloc As Integer = 400
    Dim yloc As Integer = 300
    Dim xlocb As Integer
    Dim ylocb As Integer

    Dim pointCount As Integer = 1
    Dim pointsX(1) As Single
    Dim pointsY(1) As Single
    Dim pointsZ(1) As Single

    Dim point1(1) As Integer
    Dim point2(1) As Integer

    Dim pointsXBackup() As Single
    Dim pointsYBackup() As Single
    Dim pointsZBackup() As Single

    Dim lastpoints As Point()
    Dim isOnScreen = False

    Dim lastAngleX = 0
    Dim lastAngleY = 0
    Dim lastAngleZ = 0

    Dim lastX = 0
    Dim lastY = 0

    Dim lastZ As Integer = 0

    Dim dragMode As Boolean = False

    Dim fov As Single

    Private Sub Form1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles Me.KeyPress
        If e.KeyChar = "c" Then
            If iscapture = True Then
                iscapture = False
                Windows.Forms.Cursor.Show()
            End If
        End If
    End Sub

    Sub setDefault()
        pointsX = pointsXBackup.Clone
        pointsY = pointsYBackup.Clone
        pointsZ = pointsZBackup.Clone

        TrackBar2.Value = 90
        TrackBar3.Value = 0
        TrackBar4.Value = 0
        TrackBar5.Value = 0
        TrackBar6.Value = 50

        fov = Math.PI * (TrackBar2.Value / 180)
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        fov = Math.PI * (TrackBar2.Value / 180)

    End Sub

    Function getLoc(ByVal point As Single, ByVal z As Single, ByVal screen As Integer)

        Dim isNegative As Boolean = False
        Dim screenMax As Single = 0
        Dim output As Single = 0

        If z = 0 Then
            z = -1
        End If
        If point <> 0 Then
            If point < 0 Then
                isNegative = True
                point = point * (-1)
            End If

            screenMax = Math.Tan(fov / 2) * z
            output = (point / screenMax) * (screen / 2)

            If isNegative = True Then
                output = (screen / 2) + output
            Else
                output = (screen / 2) - output
            End If

        Else
            output = screen / 2
        End If

        If z < 0 Then
            output *= -1
        End If

        Return output
    End Function

    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        Me.CreateGraphics.Clear(Color.White)
        render()
    End Sub

    Sub render()
        If TrackBar7.Value = 0 Then
            renderBlack()
        Else
            renderRed()
            renderBlue()
        End If

        'Dim i As Integer
        'Dim dots(pointCount - 1) As PointF

        'For i = 0 To pointCount - 1
        '    Dim point As New PointF(pointsX(i), pointsY(i))

        '    point.X = getLoc(point.X, TrackBar1.Value + pointsZ(i), Me.Width - 16)
        '    point.Y = getLoc(point.Y, TrackBar1.Value + pointsZ(i), Me.Height - 39)

        '    dots(i) = point

        'Next

        'Dim j As Integer

        'For i = 0 To pointCount - 2
        '    For j = i + 1 To pointCount - 1
        '        Me.CreateGraphics.DrawLine(Pens.Red, dots(i).X, dots(i).Y, dots(j).X, dots(j).Y)
        '    Next
        'Next

    End Sub
    Sub renderBlack()

        Dim i As Integer
        Dim dots(pointCount - 1) As PointF

        For i = 0 To pointCount - 1
            Dim point As New PointF(pointsX(i), pointsY(i))

            point.X -= TrackBar9.Value
            point.Y += TrackBar8.Value

            point.X = getLoc(point.X, TrackBar1.Value + pointsZ(i), Fwidth) + ((Me.Width - Fwidth) / 2)
            point.Y = getLoc(point.Y, TrackBar1.Value + pointsZ(i), Fheight)

            dots(i) = point

        Next

        If CheckBox1.Checked = True Then
            Dim j As Integer
            For i = 0 To pointCount - 2
                For j = i + 1 To pointCount - 1
                    Me.CreateGraphics.DrawLine(Pens.Black, dots(i).X, dots(i).Y, dots(j).X, dots(j).Y)
                Next
            Next
        Else
            For i = 0 To point1.Length - 1
                Me.CreateGraphics.DrawLine(Pens.Black, dots(point1(i)).X, dots(point1(i)).Y, dots(point2(i)).X, dots(point2(i)).Y)
            Next
        End If

    End Sub
    Sub renderRed()

        Dim i As Integer
        Dim dots(pointCount - 1) As PointF

        For i = 0 To pointCount - 1
            Dim point As New PointF(pointsX(i), pointsY(i))

            point.X -= TrackBar9.Value
            point.Y += TrackBar8.Value

            point.X = point.X - (TrackBar7.Value / 1000)

            point.X = getLoc(point.X, TrackBar1.Value + pointsZ(i), Fwidth) + ((Me.Width - Fwidth) / 2)
            point.Y = getLoc(point.Y, TrackBar1.Value + pointsZ(i), Fheight)

            dots(i) = point

        Next

        If CheckBox1.Checked = True Then
            Dim j As Integer
            For i = 0 To pointCount - 2
                For j = i + 1 To pointCount - 1
                    Me.CreateGraphics.DrawLine(Pens.Coral, dots(i).X, dots(i).Y, dots(j).X, dots(j).Y)
                Next
            Next
        Else
            For i = 0 To point1.Length - 1
                Me.CreateGraphics.DrawLine(Pens.Coral, dots(point1(i)).X, dots(point1(i)).Y, dots(point2(i)).X, dots(point2(i)).Y)
            Next
        End If

    End Sub

    Sub renderBlue()

        Dim i As Integer
        Dim dots(pointCount - 1) As PointF

        For i = 0 To pointCount - 1
            Dim point As New PointF(pointsX(i), pointsY(i))

            point.X -= TrackBar9.Value
            point.Y += TrackBar8.Value

            point.X = point.X + (TrackBar7.Value / 1000)

            point.X = getLoc(point.X, TrackBar1.Value + pointsZ(i), Fwidth) + ((Me.Width - Fwidth) / 2)
            point.Y = getLoc(point.Y, TrackBar1.Value + pointsZ(i), Fheight)

            dots(i) = point

        Next

        If CheckBox1.Checked = True Then
            Dim j As Integer
            For i = 0 To pointCount - 2
                For j = i + 1 To pointCount - 1
                    Me.CreateGraphics.DrawLine(Pens.Aqua, dots(i).X, dots(i).Y, dots(j).X, dots(j).Y)
                Next
            Next
        Else
            For i = 0 To point1.Length - 1
                Me.CreateGraphics.DrawLine(Pens.Aqua, dots(point1(i)).X, dots(point1(i)).Y, dots(point2(i)).X, dots(point2(i)).Y)
            Next
        End If

    End Sub

    Private Sub Form1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseDown
        dragMode = True
        lastX = e.Y
        lastY = e.X
    End Sub

    Private Sub Form1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseMove

        If dragMode = True Then
            Me.CreateGraphics.Clear(Color.White)
            Dim angleX As Single = ((lastX - e.Y) / (Me.Width - 16)) * 360 * (TrackBar6.Value / 50)
            Dim angleY As Single = ((lastY - e.X) / (Me.Height - 39)) * 360 * (TrackBar6.Value / 50)
            Label1.Text = e.Y
            Label2.Text = e.X

            rotate(angleX, "x")
            rotate(angleY, "y")

            render()

            lastX = e.Y
            lastY = e.X


        End If
    End Sub

    Private Sub Form1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Me.MouseUp
        dragMode = False
    End Sub

    Private Sub Form1_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        Me.CreateGraphics.Clear(Color.White)
        render()
    End Sub

    Sub rotate(ByVal angle As Single, ByVal dir As Char)
        angle = Math.PI * (angle / 180)

        Dim rotationMatrix(3, 3) As Single
        If dir = "y" Then
            rotationMatrix(0, 0) = Math.Cos(angle)
            rotationMatrix(0, 1) = 0
            rotationMatrix(0, 2) = -Math.Sin(angle)

            rotationMatrix(1, 0) = 0
            rotationMatrix(1, 1) = 1
            rotationMatrix(1, 2) = 0

            rotationMatrix(2, 0) = Math.Sin(angle)
            rotationMatrix(2, 1) = 0
            rotationMatrix(2, 2) = Math.Cos(angle)

        ElseIf dir = "x" Then

            rotationMatrix(0, 0) = 1
            rotationMatrix(0, 1) = 0
            rotationMatrix(0, 2) = 0

            rotationMatrix(1, 0) = 0
            rotationMatrix(1, 1) = Math.Cos(angle)
            rotationMatrix(1, 2) = -Math.Sin(angle)

            rotationMatrix(2, 0) = 0
            rotationMatrix(2, 1) = Math.Sin(angle)
            rotationMatrix(2, 2) = Math.Cos(angle)

        ElseIf dir = "z" Then

            rotationMatrix(0, 0) = Math.Cos(angle)
            rotationMatrix(0, 1) = -Math.Sin(angle)
            rotationMatrix(0, 2) = 0

            rotationMatrix(1, 0) = Math.Sin(angle)
            rotationMatrix(1, 1) = Math.Cos(angle)
            rotationMatrix(1, 2) = 0

            rotationMatrix(2, 0) = 0
            rotationMatrix(2, 1) = 0
            rotationMatrix(2, 2) = 1

        End If

        matrixMultiply(rotationMatrix)

    End Sub

    Sub matrixMultiply(ByRef rotationMatrix(,) As Single)
        Dim i As Integer = 0
        Dim j As Integer = 0
        Dim x As Single
        Dim y As Single
        Dim z As Single
        For j = 0 To pointCount - 1
            x = (rotationMatrix(0, 0) * pointsX(j)) + (rotationMatrix(0, 1) * pointsY(j)) + (rotationMatrix(0, 2) * pointsZ(j))
            y = (rotationMatrix(1, 0) * pointsX(j)) + (rotationMatrix(1, 1) * pointsY(j)) + (rotationMatrix(1, 2) * pointsZ(j))
            z = (rotationMatrix(2, 0) * pointsX(j)) + (rotationMatrix(2, 1) * pointsY(j)) + (rotationMatrix(2, 2) * pointsZ(j))

            pointsX(j) = x
            pointsY(j) = y
            pointsZ(j) = z

        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Me.CreateGraphics.Clear(Color.White)
        Dim s As Single = 2.5
        If RadioButton1.Checked Then
            rotate(s, "x")
        ElseIf RadioButton2.Checked Then
            rotate(s, "y")
        ElseIf RadioButton3.Checked Then
            rotate(s, "z")
        End If
        render()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Timer1.Enabled = False
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Me.CreateGraphics.Clear(Color.White)
        setDefault()
        render()
    End Sub

    Private Sub TrackBar2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar2.Scroll
        Label11.Text = TrackBar2.Value
        fov = Math.PI * (TrackBar2.Value / 180)
        Me.CreateGraphics.Clear(Color.White)
        render()
    End Sub


    Private Sub TrackBar3_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar3.Scroll
        Me.CreateGraphics.Clear(Color.White)
        rotate(lastAngleY - TrackBar3.Value, "y")
        lastAngleY = TrackBar3.Value
        render()
    End Sub

    Private Sub TrackBar4_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar4.Scroll
        Me.CreateGraphics.Clear(Color.White)
        rotate(TrackBar4.Value - lastAngleX, "x")
        lastAngleX = TrackBar4.Value
        render()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "" Then
            Button4.Enabled = False
            Dim input() As Char
            Dim xyz As Integer = 0
            input = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName).ToCharArray
            Dim tmp As String = ""
            Dim i As Integer
            For i = 0 To input.Length - 1

                If input(i) = "," Then

                    If xyz = 0 Then
                        pointsX(pointCount - 1) = Val(tmp)
                    ElseIf xyz = 1 Then
                        pointsY(pointCount - 1) = Val(tmp)
                    Else
                        pointsZ(pointCount - 1) = Val(tmp)
                    End If
                    xyz += 1
                    If xyz > 2 Then
                        xyz = 0

                        Array.Resize(pointsX, pointCount + 1)
                        Array.Resize(pointsY, pointCount + 1)
                        Array.Resize(pointsZ, pointCount + 1)
                        pointCount += 1
                    End If
                    tmp = ""

                ElseIf input(i) = Chr(13) Then

                Else
                    tmp = tmp & input(i)
                End If

            Next
            pointsZ(pointCount - 1) = Val(tmp)
        End If
        render()
    End Sub

    Private Sub TrackBar5_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar5.Scroll
        Me.CreateGraphics.Clear(Color.White)
        rotate(lastAngleZ - TrackBar5.Value, "z")
        lastAngleZ = TrackBar5.Value
        render()
    End Sub

    Private Sub Form1_Shown(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Shown
        OpenFileDialog1.FileName = Nothing
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> Nothing Then
            Button4.Enabled = False
            Dim input() As Char
            Dim xyz As Integer = 0
            Dim mode As Integer = 0
            input = My.Computer.FileSystem.ReadAllText(OpenFileDialog1.FileName).ToCharArray
            Dim tmp As String = ""
            Dim i As Integer
            For i = 0 To input.Length - 1

                If input(i) = "," Then
                    If mode = 0 Then
                        If xyz = 0 Then
                            pointsX(pointCount - 1) = Val(tmp)
                        ElseIf xyz = 1 Then
                            pointsY(pointCount - 1) = Val(tmp)
                        Else
                            pointsZ(pointCount - 1) = Val(tmp)
                        End If
                        xyz += 1
                        If xyz > 2 Then
                            xyz = 0

                            Array.Resize(pointsX, pointCount + 1)
                            Array.Resize(pointsY, pointCount + 1)
                            Array.Resize(pointsZ, pointCount + 1)
                            pointCount += 1
                        End If
                        tmp = ""
                    ElseIf mode = 1 Then
                        point1(point1.Length - 1) = Val(tmp)
                        Array.Resize(point1, point1.Length + 1)
                        tmp = ""
                    ElseIf mode = 2 Then
                        point2(point2.Length - 1) = Val(tmp)
                        Array.Resize(point2, point2.Length + 1)
                        tmp = ""
                    End If

                ElseIf input(i) = Chr(10) Then
                    If mode = 0 Then
                        pointsZ(pointCount - 1) = Val(tmp)
                        tmp = ""
                        mode = 1
                    ElseIf mode = 1 Then
                        point1(point1.Length - 1) = Val(tmp)
                        tmp = ""
                        mode = 2
                    End If
                Else
                    tmp = tmp & input(i)
                End If

            Next
            point2(point2.Length - 1) = Val(tmp)

            pointsXBackup = pointsX.Clone
            pointsYBackup = pointsY.Clone
            pointsZBackup = pointsZ.Clone

        End If

        Select Case ComboBox1.SelectedIndex
            Case 0
                Fheight = Me.Height - 39
                Fwidth = Fheight

            Case 1
                Fheight = Me.Height - 39
                Fwidth = (Fheight / 3) * 4

            Case 2
                Fheight = Me.Height - 39
                Fwidth = (Fheight / 9) * 16

        End Select

        render()
    End Sub

    Private Sub TrackBar7_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar7.Scroll
        Me.CreateGraphics.Clear(Color.White)
        render()
    End Sub

    Private Sub CheckBox1_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        Me.CreateGraphics.Clear(Color.White)
        render()
    End Sub

    Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
        Me.CreateGraphics.Clear(Color.White)
        If CheckBox2.Checked = True Then
            Dim i As Integer = 0
            For i = 0 To pointCount - 1
                pointsZ(i) += TrackBar1.Value
            Next
        Else
            setDefault()
        End If
        render()
    End Sub

    Private Sub TrackBar9_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar9.Scroll
        Me.CreateGraphics.Clear(Color.White)
        render()
    End Sub

    Private Sub TrackBar8_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar8.Scroll
        Me.CreateGraphics.Clear(Color.White)
        render()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MsgBox(ComboBox1.SelectedIndex)
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedIndex
            Case 0
                Fheight = Me.Height - 39
                Fwidth = Fheight

            Case 1
                Fheight = Me.Height - 39
                Fwidth = (Fheight / 3) * 4

            Case 2
                Fheight = Me.Height - 39
                Fwidth = (Fheight / 9) * 16

        End Select
        Me.CreateGraphics.Clear(Color.White)
        render()
    End Sub

    Private Sub Button5_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        bitmap = My.Computer.FileSystem.ReadAllBytes("template.bmp")

        Dim i As Integer
        Dim dots(pointCount - 1) As PointF

        For i = 0 To pointCount - 1
            Dim point As New PointF(pointsX(i), pointsY(i))

            point.X -= TrackBar9.Value
            point.Y += TrackBar8.Value

            point.X = getLoc(point.X, TrackBar1.Value + pointsZ(i), Fwidth) + ((Me.Width - Fwidth) / 2)
            point.Y = getLoc(point.Y, TrackBar1.Value + pointsZ(i), Fheight)

            dots(i) = point

        Next
        
        For i = 0 To point1.Length - 1
            drawLine(dots(point1(i)).X, dots(point1(i)).Y, dots(point2(i)).X, dots(point2(i)).Y)
        Next

        My.Computer.FileSystem.WriteAllBytes("export.bmp", bitmap, False)

    End Sub

    Sub drawLine(ByVal x1 As Integer, ByVal y1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer)
        Dim x As Integer
        Dim y As Integer

        If x1 > x2 Then
            For x = x2 To x1
                If x1 <> x2 Then
                    y = y1 + (((y1 - y2) / (x1 - x2)) * (x - x1))
                    If (x >= 0) And (x < 1280) And (y >= 0) And (y < 720) Then
                        bitmap(bitmap.Length - 1281 + x - (y * 1280)) = 0
                    End If
                End If
            Next
        Else
            For x = x1 To x2
                If x1 <> x2 Then
                    y = y1 + (((y1 - y2) / (x1 - x2)) * (x - x1))
                    If (x >= 0) And (x < 1280) And (y >= 0) And (y < 720) Then
                        bitmap(bitmap.Length - 1281 + x - (y * 1280)) = 0
                    End If
                End If
            Next
        End If

        If y1 > y2 Then
            For y = y2 To y1
                If y1 <> y2 Then
                    x = (((y - y1) / (y1 - y2)) * (x1 - x2)) + x1
                    If (x >= 0) And (x < 1280) And (y >= 0) And (y < 720) Then
                        bitmap(bitmap.Length - 1281 + x - (y * 1280)) = 0
                    End If
                End If
            Next
        Else
            For y = y1 To y2
                If y1 <> y2 Then
                    x = (((y - y1) / (y1 - y2)) * (x1 - x2)) + x1
                    If (x >= 0) And (x < 1280) And (y >= 0) And (y < 720) Then
                        bitmap(bitmap.Length - 1281 + x - (y * 1280)) = 0
                    End If
                End If
            Next
        End If

    End Sub

End Class
