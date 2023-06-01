Public Class Form1
    Dim CurrentGoal As Integer
    Dim Speeds As Double() = {
        1.0,
        2.0,
        3.0,
        5.0
    }
    Dim Timeouts As Double() = {
        10.0,
        7.0,
        4.0,
        3.0
    }
    Dim GameRunning As Boolean = False
    Dim Score As Integer

    Enum FireflyState
        Wait
        Flying
        Escaping
        Gone
    End Enum

    Structure Firefly
        Dim Position As PointF
        Dim Destination As PointF
        Dim Color As Color
        Dim Speed As Double
        Dim State As FireflyState
        Dim Timeout As Double
    End Structure

    Dim Flies As Firefly()
    Dim FlyIndex As Integer
    Dim FrameInterval As Integer
    Dim CurrentSpeed As Double

    Dim GameStartTime As Date
    Dim GameElapsedTime As TimeSpan
    Dim LastFrameTime As Date

    Dim BushCurves1(14) As PointF
    Dim BushCurves2(14) As PointF

    Dim Stars(199) As PointF

    Private Sub NewGameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewGameToolStripMenuItem.Click
        Form2.Show()
    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        Close()
    End Sub

    Public Sub StartNewGame(Goal As Integer, Speed As Integer)
        If Speed = -1 Then
            MsgBox("Error: Speed is invalid", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Start new game")
            Exit Sub
        End If
        ReDim Flies(Goal - 1)
        CurrentGoal = Goal
        FlyIndex = 0

        ' Initialize every fly
        Dim FlyRandom = New Random()
        Dim FlyRandomRedColor = New Random()
        For i = 0 To Flies.Length - 1
            Flies(i).Position = New PointF(FlyRandom.NextSingle, 1)
            Flies(i).Color = Color.FromArgb(128, 255, CInt(FlyRandomRedColor.NextSingle * 255), 0)
            Flies(i).State = FireflyState.Wait
            Dim FlyDestX = FlyRandom.NextSingle() * 0.9 + 0.05
            Dim FlyDestY = FlyRandom.NextSingle() * 0.75 + 0.05
            Flies(i).Destination = New PointF(FlyDestX, FlyDestY)
            Flies(i).Timeout = Timeouts(Speed)
        Next
        FrameInterval = 0
        CurrentSpeed = Speeds(Speed)

        ' Generate bush
        For i = 0 To 14
            Dim R_X1 As Single
            Dim R_Y1 As Single
            Dim R_X2 As Single
            Dim R_Y2 As Single
            Dim R_Gen As New Random()
            Select Case i
                Case 0
                    R_X1 = 0
                    R_X2 = 0
                    R_Y1 = 1
                    R_Y2 = 1
                Case 1
                    R_X1 = 0
                    R_X2 = 0
                    R_Y1 = 0.83 + (R_Gen.NextSingle * 0.09)
                    R_Y2 = 0.855 + (R_Gen.NextSingle * 0.07)
                Case 2 To 12
                    R_X1 = 0.09 + ((R_Gen.NextSingle - 0.5) * 0.06) + ((i - 2) * 0.084)
                    R_X2 = 0.09 + ((R_Gen.NextSingle - 0.5) * 0.06) + ((i - 2) * 0.084)
                    R_Y1 = 0.83 + (R_Gen.NextSingle * 0.09)
                    R_Y2 = 0.855 + (R_Gen.NextSingle * 0.07)
                Case 13
                    R_X1 = 1
                    R_X2 = 1
                    R_Y1 = 0.83 + (R_Gen.NextSingle * 0.09)
                    R_Y2 = 0.855 + (R_Gen.NextSingle * 0.07)
                Case 14
                    R_X1 = 1
                    R_X2 = 1
                    R_Y1 = 1
                    R_Y2 = 1
            End Select
            BushCurves1(i) = New PointF(R_X1, R_Y1)
            BushCurves2(i) = New PointF(R_X2, R_Y2)
        Next

        ' Generate stars
        For i = 0 To 199
            Dim R_Gen As New Random()
            Stars(i) = New PointF(R_Gen.NextSingle() * 2 - 1, R_Gen.NextSingle() * 2 - 1)
        Next

        GameStartTime = Now
        LastFrameTime = Now
        Score = 0
        GameRunning = True
        Timer1.Enabled = True ' The game starts
    End Sub

    Private Sub UserControlUpdate(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim FrameElapsed As TimeSpan
        GameElapsedTime = Now.Subtract(GameStartTime)
        FrameElapsed = Now.Subtract(LastFrameTime)
        LastFrameTime = Now
        Dim fTheta As Double = GameElapsedTime.TotalSeconds
        Dim fDelta As Double = FrameElapsed.TotalSeconds
        Dim AspectRatio = PictureBox1.Width / PictureBox1.Height

        Dim Rand As New Random()

        FrameInterval -= 1
        Select Case FrameInterval
            Case -1
                FrameInterval = 60.0 / CurrentSpeed
            Case 0
                If FlyIndex = CurrentGoal Then
                    Exit Select
                End If
                Flies(FlyIndex).State = FireflyState.Flying
                Flies(FlyIndex).Speed = CurrentSpeed
                If FlyIndex < CurrentGoal Then
                    FlyIndex += 1
                End If
        End Select

        For i = 0 To CurrentGoal - 1
            Select Case Flies(i).State
                Case FireflyState.Wait, FireflyState.Gone
                    Continue For
                Case FireflyState.Flying, FireflyState.Escaping
                    Dim FlySpeed = Flies(i).Speed * 0.1
                    If Flies(i).State = FireflyState.Escaping Then
                        FlySpeed *= 2
                    End If
                    Dim FlyPosX = Flies(i).Position.X
                    Dim FlyPosY = Flies(i).Position.Y
                    Dim FlyDestX = Flies(i).Destination.X
                    Dim FlyDestY = Flies(i).Destination.Y
                    Dim XDiff = FlyDestX - FlyPosX
                    Dim YDiff = FlyDestY - FlyPosY
                    Dim NewDest As Boolean = False ' If the fly reached its destination
                    If Math.Abs(XDiff) > Math.Abs(YDiff) Then
                        If XDiff > 0 Then
                            FlyPosX += FlySpeed * fDelta
                            If YDiff > 0 Then
                                FlyPosY -= FlySpeed * fDelta * (YDiff / XDiff)
                            Else
                                FlyPosY += FlySpeed * fDelta * (YDiff / XDiff)
                            End If
                        Else
                            FlyPosX -= FlySpeed * fDelta
                            If YDiff > 0 Then
                                FlyPosY -= FlySpeed * fDelta * (YDiff / XDiff)
                            Else
                                FlyPosY += FlySpeed * fDelta * (YDiff / XDiff)
                            End If
                        End If
                        If Math.Abs(XDiff) < FlySpeed * fDelta Then
                            FlyPosX = FlyDestX
                            FlyPosY = FlyDestY
                            NewDest = True
                        End If
                    Else
                        If YDiff > 0 Then
                            FlyPosY += FlySpeed * fDelta
                            If XDiff > 0 Then
                                FlyPosX += FlySpeed * fDelta * (XDiff / YDiff)
                            Else
                                FlyPosX -= FlySpeed * fDelta * (XDiff / YDiff)
                            End If
                        Else
                            FlyPosY -= FlySpeed * fDelta
                            If XDiff > 0 Then
                                FlyPosX += FlySpeed * fDelta * (XDiff / YDiff)
                            Else
                                FlyPosX -= FlySpeed * fDelta * (XDiff / YDiff)
                            End If
                        End If
                        If Math.Abs(YDiff) < FlySpeed * fDelta Then
                            FlyPosX = FlyDestX
                            FlyPosY = FlyDestY
                            NewDest = True
                        End If
                    End If
                    If NewDest = True Then
                        If Flies(i).State = FireflyState.Escaping Then
                            Flies(i).State = FireflyState.Gone
                            Exit For
                        End If
                        FlyDestX = Rand.NextSingle() * 0.9 + 0.05
                        FlyDestY = Rand.NextSingle() * 0.75 + 0.05
                        Flies(i).Destination = New PointF(FlyDestX, FlyDestY)
                    End If
                    Flies(i).Position = New PointF(FlyPosX, FlyPosY)
                    Flies(i).Timeout -= fDelta
                    If Flies(i).Timeout < 0 And Flies(i).State = FireflyState.Flying Then
                        Flies(i).State = FireflyState.Escaping
                        FlyDestX = 1.5 * Math.Pow(-1, Rand.NextInt64() Mod 2)
                        FlyDestY = Rand.NextSingle() * 1.5 - 1.0
                        Flies(i).Destination = New PointF(FlyDestX, FlyDestY)
                    End If
            End Select
        Next

        If PictureBox1.Height <= 0 Then
            Exit Sub
        End If

        Dim BackBuffer As New Bitmap(PictureBox1.Width, PictureBox1.Height)
        Dim g As Graphics
        g = Graphics.FromImage(BackBuffer)
        g.Clear(Color.FromArgb(255, 0, 0, 32)) ' Background

        ' DRAW DEBUG INFO
#If DEBUG Then
        Dim DebugBrush As New SolidBrush(Color.Teal)
        g.DrawString("FPS: " + CStr(1 / fDelta), Font, DebugBrush, New PointF(10, 10))
        g.DrawString("FPSInv: " + CStr(fDelta), Font, DebugBrush, New PointF(10, 20))
        g.DrawString("Fly0 Pos: " + CStr(Flies(0).Position.X) + ", " + CStr(Flies(0).Position.Y), Font, DebugBrush, New PointF(10, 30))
        g.DrawString("Fly0 Destination: " + CStr(Flies(0).Destination.X) + ", " + CStr(Flies(0).Destination.Y), Font, DebugBrush, New PointF(10, 40))
        g.DrawString("Fly0 State: " + CStr(Flies(0).State), Font, DebugBrush, New PointF(10, 50))
#End If

        ' (maybe) draw rotating stars, decors like trees...
        For i = 0 To 199
            Dim S_X = (0.5 + (Stars(i).X * Math.Cos(fTheta / 7) + Stars(i).Y * -Math.Sin(fTheta / 7))) * PictureBox1.Width
            Dim S_Y = (0.5 + ((Stars(i).X * Math.Sin(fTheta / 7) + Stars(i).Y * Math.Cos(fTheta / 7))) * AspectRatio) * PictureBox1.Height
            g.DrawImage(My.Resources.Star, New PointF(S_X, S_Y))
        Next

        ' DRAW FLIES
        For i = 0 To Flies.Length - 1
            If Flies(i).State = FireflyState.Gone Then
                Continue For
            End If
            g.FillEllipse(New SolidBrush(Flies(i).Color), Flies(i).Position.X * PictureBox1.Width, Flies(i).Position.Y * PictureBox1.Height, CInt(0.05 * PictureBox1.Width), CInt(0.05 * PictureBox1.Height * AspectRatio))
#If DEBUG Then
            Dim DebugBrush_Pos As New SolidBrush(Color.SeaGreen)
            Dim DebugBrush_Dest As New SolidBrush(Color.OrangeRed)
            g.DrawString(CStr(i), Font, DebugBrush_Pos, New PointF(Flies(i).Position.X * PictureBox1.Width, Flies(i).Position.Y * PictureBox1.Height))
            g.DrawString(CStr(i), Font, DebugBrush_Dest, New PointF(Flies(i).Destination.X * PictureBox1.Width, Flies(i).Destination.Y * PictureBox1.Height))
#End If
        Next

        ' DRAW BUSH
        Dim ScaledBushCurves1 As PointF() = BushCurves1.Clone()
        Dim ScaledBushCurves2 As PointF() = BushCurves2.Clone()
        For i = 0 To 14
            ScaledBushCurves1(i) = New PointF(ScaledBushCurves1(i).X * PictureBox1.Width, ScaledBushCurves1(i).Y * PictureBox1.Height)
            ScaledBushCurves2(i) = New PointF(ScaledBushCurves2(i).X * PictureBox1.Width, ScaledBushCurves2(i).Y * PictureBox1.Height)
        Next
        g.FillClosedCurve(New SolidBrush(Color.DarkOliveGreen), ScaledBushCurves2)
        g.FillClosedCurve(New SolidBrush(Color.ForestGreen), ScaledBushCurves1)

        ' DRAW SCORE
        g.DrawString("Score: " + CStr(Score), Font, New SolidBrush(Color.Honeydew), New PointF(4, PictureBox1.Height - 20))

        PictureBox1.Image = BackBuffer
        Invalidate()
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As MouseEventArgs) Handles PictureBox1.Click
        Dim AspectRatio = PictureBox1.Width / PictureBox1.Height
        If GameRunning = False Then
            Exit Sub
        End If
        For i = 0 To Flies.Length - 1
            If Flies(i).State <> FireflyState.Gone And e.X > (Flies(i).Position.X * PictureBox1.Width) And e.X < ((Flies(i).Position.X * PictureBox1.Width) + CInt(0.05 * PictureBox1.Width)) And e.Y > (Flies(i).Position.Y * PictureBox1.Height) And e.Y < ((Flies(i).Position.Y * PictureBox1.Height) + CInt(0.05 * PictureBox1.Height * AspectRatio)) Then
                Flies(i).State = FireflyState.Gone
                Score += 1
                Exit Sub
            End If
        Next
    End Sub
End Class
