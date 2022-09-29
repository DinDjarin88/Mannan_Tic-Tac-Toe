Public Class Form1
    Dim Button_Click As Integer
    Dim Player_1 As String
    Dim Player_2 As String
    Dim P1pts = 0
    Dim P2pts = 0
    Dim Tie_Counter = 0
    Dim Turn_Count As Integer
    Dim Reset_Board As Integer
    Dim ResetNum As Integer
    Dim ExitConfirm
    Dim ResetConfirm

    Public Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Player_1 = InputBox("Input Player 1 Name")

        Player_2 = InputBox("Input Player 2 Name")

        Label5.Text = Player_1 + ("'s points:")

        Label6.Text = Player_2 + ("'s points:")

        Button_Click = 1
        Turn_Count = 0
        TextBox6.Enabled = False
    End Sub


    Private Sub Button(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Btn4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click

        If Button_Click = 1 Then
            DirectCast(sender, Button).Text = "X"
            DirectCast(sender, Button).Enabled = False
            Button_Click -= 1
            Turn_Count += 1

        Else
            DirectCast(sender, Button).Text = "O"
            DirectCast(sender, Button).Enabled = False
            Button_Click += 1
            Turn_Count += 1

        End If

        If Button1.Text = "X" And Button2.Text = "X" And Button3.Text = "X" Or
                    Button1.Text = "X" And Button5.Text = "X" And Button9.Text = "X" Or
                    Btn4.Text = "X" And Button5.Text = "X" And Button6.Text = "X" Or
                    Button7.Text = "X" And Button8.Text = "X" And Button9.Text = "X" Or
                    Button1.Text = "X" And Btn4.Text = "X" And Button7.Text = "X" Or
                    Button2.Text = "X" And Button5.Text = "X" And Button8.Text = "X" Or
                    Button3.Text = "X" And Button6.Text = "X" And Button9.Text = "X" Or
                    Button3.Text = "X" And Button5.Text = "X" And Button7.Text = "X" Then
            P1pts += 1
            Turn_Count = 0
            Label2.Text = P1pts
            MsgBox(Player_1 + " Wins!")
            Reset_Board += 1

        End If




        If Button1.Text = "O" And Button2.Text = "O" And Button3.Text = "O" Or Button1.Text = "O" And Button5.Text = "O" And Button9.Text = "O" Or
            Btn4.Text = "O" And Button5.Text = "O" And Button6.Text = "O" Or
            Button3.Text = "O" And Button6.Text = "O" And Button9.Text = "O" Or
            Button7.Text = "O" And Button8.Text = "O" And Button9.Text = "O" Or
            Button1.Text = "O" And Btn4.Text = "O" And Button7.Text = "O" Or
            Button2.Text = "O" And Button5.Text = "O" And Button8.Text = "O" Or
            Button3.Text = "O" And Button5.Text = "O" And Button7.Text = "O" Then

            MessageBox.Show(Player_2 + " Wins!")
            P2pts += 1
            Label3.Text = P2pts
            Reset_Board += 1

        End If

        If Reset_Board = 1 Then
            Button1.Text = ""
            Button2.Text = ""
            Button3.Text = ""
            Btn4.Text = ""
            Button5.Text = ""
            Button6.Text = ""
            Button7.Text = ""
            Button8.Text = ""
            Button9.Text = ""
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Btn4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
            Button7.Enabled = True
            Button8.Enabled = True
            Button9.Enabled = True
            Turn_Count = 0
            Reset_Board = 0
        End If


        If Turn_Count = 9 Then
            MessageBox.Show("IT'S A DRAW!!!!")
            Button1.Text = ""
            Button2.Text = ""
            Button3.Text = ""
            Btn4.Text = ""
            Button5.Text = ""
            Button6.Text = ""
            Button7.Text = ""
            Button8.Text = ""
            Button9.Text = ""
            Button1.Enabled = True
            Button2.Enabled = True
            Button3.Enabled = True
            Btn4.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
            Button7.Enabled = True
            Button8.Enabled = True
            Button9.Enabled = True
            Tie_Counter += 1
            Label4.Text = Tie_Counter

        End If



    End Sub
    Private Sub Reset(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click, Button3.Click, Btn4.Click, Button5.Click, Button6.Click, Button7.Click, Button8.Click, Button9.Click, ResetButton.Click
        If sender Is ResetButton Then
            ResetConfirm = MsgBox("Are You Sure?", (vbYesNo + vbQuestion))
            If ResetConfirm = DialogResult.Yes Then
                Turn_Count = 0
                Button1.Text = ""
                Button2.Text = ""
                Button3.Text = ""
                Btn4.Text = ""
                Button5.Text = ""
                Button6.Text = ""
                Button7.Text = ""
                Button8.Text = ""
                Button9.Text = ""
                Button1.Enabled = True
                Button2.Enabled = True
                Button3.Enabled = True
                Btn4.Enabled = True
                Button5.Enabled = True
                Button6.Enabled = True
                Button7.Enabled = True
                Button8.Enabled = True
                Button9.Enabled = True
                Label2.Text = 0
                Label3.Text = 0
                Label4.Text = 0
                P2pts = 0
                P1pts = 0
            End If
        End If

    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        ExitConfirm = MsgBox("Are You Sure?", (vbYesNo + vbQuestion))
        If ExitConfirm = DialogResult.Yes Then
            Me.Close()
        End If

    End Sub


    Private Sub Rename(sender As Object, e As EventArgs) Handles Rename_Button.Click
        Dim Player_1 = InputBox("Input Player 1 Name")

        Dim Player_2 = InputBox("Input Player 2 Name")

        Label5.Text = Player_1 + ("'s points:")
        Label6.Text = Player_2 + ("'s points:")
    End Sub


End Class

