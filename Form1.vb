Public Class Form1
    Dim Unit1 As String
    Dim Unit2 As String
    Dim Unit3 As String
    Dim Unit6 As String
    Dim Unit7 As String
    Dim Unit9 As String
    Dim Unit10 As String
    Dim Unit11 As String
    Dim Unit14 As String
    Dim Unit15 As String
    Dim Unit17 As String
    Dim Unit18 As String
    Dim Unit20 As String
    Dim Unit22 As String
    Dim Unit23 As String
    Dim Unit27 As String
    Dim Unit28 As String
    Dim Unit29 As String
    Dim Sum As String
    Dim iExit As DialogResult
    Private Sub CmBox14_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CmBox14.SelectedIndexChanged

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Unit1 = CmBox1.SelectedItem
        Unit2 = CmBox2.SelectedItem
        Unit3 = CmBox3.SelectedItem
        Unit6 = CmBox4.SelectedItem
        Unit7 = CmBox5.SelectedItem
        Unit9 = CmBox7.SelectedItem
        Unit10 = CmBox6.SelectedItem
        Unit11 = CmBox8.SelectedItem
        Unit14 = CmBox9.SelectedItem
        Unit15 = CmBox13.SelectedItem
        Unit17 = CmBox11.SelectedItem
        Unit18 = CmBox15.SelectedItem
        Unit20 = CmBox10.SelectedItem
        Unit22 = CmBox14.SelectedItem
        Unit23 = CmBox12.SelectedItem
        Unit27 = CmBox16.SelectedItem
        Unit28 = CmBox17.SelectedItem
        Unit29 = CmBox18.SelectedItem

        If CmBox1.SelectedItem = "P" Then
            Unit1 = 70
        ElseIf CmBox1.SelectedItem = "M" Then
            Unit1 = 80
        ElseIf CmBox1.SelectedItem = "D" Then
            Unit1 = 90
        End If

        If CmBox2.SelectedItem = "P" Then
            Unit2 = 70
        ElseIf CmBox2.SelectedItem = "M" Then
            Unit2 = 80
        ElseIf CmBox2.SelectedItem = "D" Then
            Unit2 = 90
        End If

        If CmBox3.SelectedItem = "P" Then
            Unit3 = 70
        ElseIf CmBox3.SelectedItem = "M" Then
            Unit3 = 80
        ElseIf CmBox3.SelectedItem = "D" Then
            Unit3 = 90
        End If

        If CmBox4.SelectedItem = "P" Then
            Unit6 = 70
        ElseIf CmBox4.SelectedItem = "M" Then
            Unit6 = 80
        ElseIf CmBox4.SelectedItem = "D" Then
            Unit6 = 90
        End If

        If CmBox5.SelectedItem = "P" Then
            Unit7 = 70
        ElseIf CmBox5.SelectedItem = "M" Then
            Unit7 = 80
        ElseIf CmBox5.SelectedItem = "D" Then
            Unit7 = 90
        End If

        If CmBox7.SelectedItem = "P" Then
            Unit9 = 70
        ElseIf CmBox7.SelectedItem = "M" Then
            Unit9 = 80
        ElseIf CmBox7.SelectedItem = "D" Then
            Unit9 = 90
        End If

        If CmBox6.SelectedItem = "P" Then
            Unit10 = 70
        ElseIf CmBox6.SelectedItem = "M" Then
            Unit10 = 80
        ElseIf CmBox6.SelectedItem = "D" Then
            Unit10 = 90
        End If

        If CmBox8.SelectedItem = "P" Then
            Unit11 = 70
        ElseIf CmBox8.SelectedItem = "M" Then
            Unit11 = 80
        ElseIf CmBox8.SelectedItem = "D" Then
            Unit11 = 90
        End If

        If CmBox9.SelectedItem = "P" Then
            Unit14 = 70
        ElseIf CmBox9.SelectedItem = "M" Then
            Unit14 = 80
        ElseIf CmBox9.SelectedItem = "D" Then
            Unit14 = 90
        End If

        If CmBox13.SelectedItem = "P" Then
            Unit15 = 70
        ElseIf CmBox13.SelectedItem = "M" Then
            Unit15 = 80
        ElseIf CmBox13.SelectedItem = "D" Then
            Unit15 = 90
        End If

        If CmBox11.SelectedItem = "P" Then
            Unit17 = 70
        ElseIf CmBox11.SelectedItem = "M" Then
            Unit17 = 80
        ElseIf CmBox11.SelectedItem = "D" Then
            Unit17 = 90
        End If

        If CmBox15.SelectedItem = "P" Then
            Unit18 = 70
        ElseIf CmBox15.SelectedItem = "M" Then
            Unit18 = 80
        ElseIf CmBox15.SelectedItem = "D" Then
            Unit18 = 90
        End If

        If CmBox10.SelectedItem = "P" Then
            Unit20 = 70
        ElseIf CmBox10.SelectedItem = "M" Then
            Unit20 = 80
        ElseIf CmBox10.SelectedItem = "D" Then
            Unit20 = 90
        End If

        If CmBox14.SelectedItem = "P" Then
            Unit22 = 70
        ElseIf CmBox14.SelectedItem = "M" Then
            Unit22 = 80
        ElseIf CmBox14.SelectedItem = "D" Then
            Unit22 = 90
        End If

        If CmBox12.SelectedItem = "P" Then
            Unit23 = 70
        ElseIf CmBox12.SelectedItem = "M" Then
            Unit23 = 80
        ElseIf CmBox12.SelectedItem = "D" Then
            Unit23 = 90
        End If

        If CmBox16.SelectedItem = "P" Then
            Unit27 = 70
        ElseIf CmBox16.SelectedItem = "M" Then
            Unit27 = 80
        ElseIf CmBox16.SelectedItem = "D" Then
            Unit27 = 90
        End If

        If CmBox17.SelectedItem = "P" Then
            Unit28 = 70
        ElseIf CmBox17.SelectedItem = "M" Then
            Unit28 = 80
        ElseIf CmBox17.SelectedItem = "D" Then
            Unit28 = 90
        End If

        If CmBox18.SelectedItem = "P" Then
            Unit29 = 70
        ElseIf CmBox18.SelectedItem = "M" Then
            Unit29 = 80
        ElseIf CmBox18.SelectedItem = "D" Then
            Unit29 = 90
        End If

        Sum = CInt(Unit1) + CInt(Unit2) + CInt(Unit3) + CInt(Unit6) + CInt(Unit7) + CInt(Unit9) + CInt(Unit10) + CInt(Unit11) + CInt(Unit14) + CInt(Unit15) + CInt(Unit17) + CInt(Unit18) + CInt(Unit20) + CInt(Unit22) + CInt(Unit23) + CInt(Unit27) + CInt(Unit28) + CInt(Unit29)
        Ans.Text = Sum

        If Ans.Text >= 1440 Then
            GradeTxt.Text = "DDD"
        ElseIf Ans.Text <= 1120 Then
            GradeTxt.Text = "PPP"
        ElseIf Ans.Text >= 1280 Then
            GradeTxt.Text = "MMM"
        End If
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit1.Text + "  :  " + CmBox1.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit2.Text + "  :  " + CmBox2.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit3.Text + "  :  " + CmBox3.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit6.Text + "  :  " + CmBox4.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit7.Text + "  :  " + CmBox5.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit9.Text + "  :  " + CmBox7.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit10.Text + "  :  " + CmBox6.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit11.Text + "  :  " + CmBox8.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit14.Text + "  :  " + CmBox9.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit15.Text + "  :  " + CmBox13.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit17.Text + "  :  " + CmBox11.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit18.Text + "  :  " + CmBox15.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit20.Text + "  :  " + CmBox10.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit22.Text + "  :  " + CmBox14.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit23.Text + "  :  " + CmBox12.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit27.Text + "  :  " + CmBox16.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit28.Text + "  :  " + CmBox17.Text + vbNewLine)
        PostTextBox.AppendText(" " + vbNewLine)
        PostTextBox.AppendText(vbTab & txtUnit29.Text + "  :  " + CmBox18.Text + vbNewLine)
        PostTextBox.AppendText("===========================================" + vbNewLine)
        PostTextBox.AppendText(vbTab & Totaltxt.Text + "  :  " + Ans.Text + vbNewLine)
        PostTextBox.AppendText("===========================================" + vbNewLine)
        PostTextBox.AppendText(vbTab & FinalGrade.Text + "  :  " + GradeTxt.Text + vbNewLine)
        PostTextBox.AppendText("===========================================" + vbNewLine)

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        iExit = MessageBox.Show("Confirm if you want to exit", "Shutting Down", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If iExit = DialogResult.Yes Then
            Application.Exit()
        End If
    End Sub
End Class
