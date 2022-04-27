Option Strict On                                     '暗黙の型変更の制限

Public Class Form1

    Private AX As Double
    Private AY As Double
    Private ASign As Integer = 0                     '式の段階を表す 0は最初、1はXまで、2は符号まで、3はYまで、4は=まで
    Private AWhichSign As Integer                    '何の符号か 1は+、2は-、3は*、4は/
    Private AS2nd As Boolean                         '次の計算に入るときのフラグ
    Private AMinus As Boolean                        '負のフラグ

    Private Sub First(ByVal TN As Integer)
        If ASign <= 1 Then
            If AMinus Then
                AX = AX * 10 - TN
                MTextBox.Text = AX.ToString
                ASign = 1
            Else
                AX = AX * 10 + TN
                MTextBox.Text = AX.ToString
                ASign = 1
            End If
        Else
            Second(TN)
        End If
    End Sub
    Private Sub SSign(ByVal TSign As Integer)
        Select Case ASign
            Case 0
                If TSign = 2 Then AMinus = True
            Case 1 To 4
                If AS2nd Then SToThe2nd()
                If ASign = 3 Then Main()
                Select Case TSign
                    Case 1
                        AWhichSign = 1
                    Case 2
                        AWhichSign = 2
                    Case 3
                        AWhichSign = 3
                    Case 4
                        AWhichSign = 4
                End Select
                ASign = 2
                AMinus = False
        End Select
    End Sub

    Private Sub Second(ByVal TN As Integer)
        If Not ASign = 3 Then AY = 0
        AY = AY * 10 + TN
        MTextBox.Text = AY.ToString
        If ASign = 4 Then AS2nd = True
        ASign = 3
    End Sub

    Private Sub Main()
        If ASign = 3 Then
            Select Case AWhichSign
                Case 1
                    AX = AX + AY
                Case 2
                    AX = AX - AY
                Case 3
                    AX = AX * AY
                Case 4
                    AX = AX / AY
            End Select
            MTextBox.Text = AX.ToString
        End If
        AS2nd = False
    End Sub
    Private Sub SC()
        AX = 0
        AY = 0
        ASign = 0
        AWhichSign = 0
        MTextBox.Text = AX.ToString
        AS2nd = False
        AMinus = False
    End Sub
    Private Sub SToThe2nd()
        Dim Temp As Double
        Temp = AY
        SC()
        AX = Temp
        MTextBox.Text = AX.ToString
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        MPanel.Visible = False
    End Sub

    Private Sub AToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AToolStripMenuItem2.Click
        MPanel.Visible = True
    End Sub

    Private Sub DToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DToolStripMenuItem1.Click
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        First(1)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        First(2)
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        First(3)
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        First(4)
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        First(5)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        First(6)
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        First(7)
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        First(8)
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        First(9)
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        First(0)
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        SSign(1)
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        SSign(2)
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        SSign(3)
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        SSign(4)
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Main()
        ASign = 4
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        SC()
    End Sub
End Class
