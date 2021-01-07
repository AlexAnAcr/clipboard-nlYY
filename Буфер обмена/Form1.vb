Public Class Form1
    Dim on_off As Boolean = False
    'All mode
    Dim mode As Integer = -1
    'Add
    Dim add_mode As Integer = -1, add_text As String = ""
    'Add repeat
    Dim add_repeat As String = ""
    'Replace
    Dim replace_a As String = "", replace_b As String = ""
    'Registr
    Dim registr As Integer = -1
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If on_off = False Then
            on_off = True
            Button1.Text = "Выключить"
            Timer1.Enabled = True
        Else
            on_off = False
            Button1.Text = "Включить"
            Timer1.Enabled = False
        End If
    End Sub
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Label2.Enabled = False
        Label3.Enabled = False
        Label4.Enabled = False
        Label5.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        ComboBox2.Enabled = False
        ComboBox3.Enabled = False
        If ComboBox1.SelectedIndex = 0 Then
            Label3.Enabled = True
            Label4.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            mode = 0
        ElseIf ComboBox1.SelectedIndex = 1 Then
            Label2.Enabled = True
            ComboBox2.Enabled = True
            TextBox1.Enabled = True
            mode = 1
        ElseIf ComboBox1.SelectedIndex = 2 Then
            Label5.Enabled = True
            ComboBox3.Enabled = True
            mode = 2
        ElseIf ComboBox1.SelectedIndex = 3 Then
            Label3.Enabled = True
            Label4.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            Label2.Enabled = True
            ComboBox2.Enabled = True
            TextBox1.Enabled = True
            mode = 3
        ElseIf ComboBox1.SelectedIndex = 4 Then
            Label2.Enabled = True
            ComboBox2.Enabled = True
            TextBox1.Enabled = True
            Label5.Enabled = True
            ComboBox3.Enabled = True
            mode = 4
        ElseIf ComboBox1.SelectedIndex = 5 Then
            Label5.Enabled = True
            ComboBox3.Enabled = True
            Label3.Enabled = True
            Label4.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            mode = 5
        ElseIf ComboBox1.SelectedIndex = 6 Then
            Label3.Enabled = True
            Label4.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            Label2.Enabled = True
            ComboBox2.Enabled = True
            TextBox1.Enabled = True
            Label5.Enabled = True
            ComboBox3.Enabled = True
            mode = 6
        End If
    End Sub
    Private Sub ComboBox3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox3.SelectedIndexChanged
        If ComboBox3.SelectedIndex = 0 Then
            registr = 0
        ElseIf ComboBox3.SelectedIndex = 1 Then
            registr = 1
        End If
    End Sub
    Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged
        If ComboBox2.SelectedIndex = 0 Then
            add_mode = 0
        ElseIf ComboBox2.SelectedIndex = 1 Then
            add_mode = 1
        ElseIf ComboBox2.SelectedIndex = 2 Then
            add_mode = 2
        End If
        add_repeat = ""
    End Sub
    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        add_text = TextBox1.Text
    End Sub
    Private Sub TextBox2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox2.TextChanged
        replace_a = TextBox2.Text
    End Sub
    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        replace_b = TextBox3.Text
    End Sub
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim times As String
        If mode = 0 Then
            If My.Computer.Clipboard.ContainsText() Then
                times = My.Computer.Clipboard.GetText()
                times = Replace(times, replace_a, replace_b)
                My.Computer.Clipboard.SetText(times)
            End If
        ElseIf mode = 1 Then
            If My.Computer.Clipboard.GetText() <> add_repeat Then
                If add_mode = 0 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = add_text & times
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                ElseIf add_mode = 1 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = times & add_text
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                ElseIf add_mode = 2 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = add_text & times & add_text
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                End If
            End If
        ElseIf mode = 2 Then
            If registr = 0 Then
                If My.Computer.Clipboard.ContainsText() Then
                    times = My.Computer.Clipboard.GetText()
                    times = (times).ToLower
                    My.Computer.Clipboard.SetText(times)
                End If
            ElseIf registr = 1 Then
                If My.Computer.Clipboard.ContainsText() Then
                    times = My.Computer.Clipboard.GetText()
                    times = (times).ToUpper
                    My.Computer.Clipboard.SetText(times)
                End If
            End If
        ElseIf mode = 3 Then
            If My.Computer.Clipboard.ContainsText() Then
                times = My.Computer.Clipboard.GetText()
                times = Replace(times, replace_a, replace_b)
                My.Computer.Clipboard.SetText(times)
            End If
            If My.Computer.Clipboard.GetText() <> add_repeat Then
                If add_mode = 0 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = add_text & times
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                ElseIf add_mode = 1 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = times & add_text
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                ElseIf add_mode = 2 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = add_text & times & add_text
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                End If
            End If
        ElseIf mode = 4 Then
            If My.Computer.Clipboard.GetText() <> add_repeat Then
                If add_mode = 0 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = add_text & times
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                ElseIf add_mode = 1 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = times & add_text
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                ElseIf add_mode = 2 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = add_text & times & add_text
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                End If
            End If
            If registr = 0 Then
                If My.Computer.Clipboard.ContainsText() Then
                    times = My.Computer.Clipboard.GetText()
                    times = (times).ToLower
                    My.Computer.Clipboard.SetText(times)
                End If
            ElseIf registr = 1 Then
                If My.Computer.Clipboard.ContainsText() Then
                    times = My.Computer.Clipboard.GetText()
                    times = (times).ToUpper
                    My.Computer.Clipboard.SetText(times)
                End If
            End If
        ElseIf mode = 5 Then
            If My.Computer.Clipboard.ContainsText() Then
                times = My.Computer.Clipboard.GetText()
                times = Replace(times, replace_a, replace_b)
                My.Computer.Clipboard.SetText(times)
            End If
            If registr = 0 Then
                If My.Computer.Clipboard.ContainsText() Then
                    times = My.Computer.Clipboard.GetText()
                    times = (times).ToLower
                    My.Computer.Clipboard.SetText(times)
                End If
            ElseIf registr = 1 Then
                If My.Computer.Clipboard.ContainsText() Then
                    times = My.Computer.Clipboard.GetText()
                    times = (times).ToUpper
                    My.Computer.Clipboard.SetText(times)
                End If
            End If
        ElseIf mode = 6 Then
            If My.Computer.Clipboard.ContainsText() Then
                times = My.Computer.Clipboard.GetText()
                times = Replace(times, replace_a, replace_b)
                My.Computer.Clipboard.SetText(times)
            End If
            If My.Computer.Clipboard.GetText() <> add_repeat Then
                If add_mode = 0 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = add_text & times
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                ElseIf add_mode = 1 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = times & add_text
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                ElseIf add_mode = 2 Then
                    If My.Computer.Clipboard.ContainsText() Then
                        times = My.Computer.Clipboard.GetText()
                        times = add_text & times & add_text
                        add_repeat = times
                        My.Computer.Clipboard.SetText(times)
                    End If
                End If
            End If
            If registr = 0 Then
                If My.Computer.Clipboard.ContainsText() Then
                    times = My.Computer.Clipboard.GetText()
                    times = (times).ToLower
                    My.Computer.Clipboard.SetText(times)
                End If
            ElseIf registr = 1 Then
                If My.Computer.Clipboard.ContainsText() Then
                    times = My.Computer.Clipboard.GetText()
                    times = (times).ToUpper
                    My.Computer.Clipboard.SetText(times)
                End If
            End If
        End If
    End Sub
End Class
