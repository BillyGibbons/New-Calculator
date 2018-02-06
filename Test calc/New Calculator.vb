Public Class Form1
    '******************************************************************
    '******New Calculator App******************************************
    '******Designed By William Gibbons 2018****************************
    '******************************************************************
    '******OP CODES****************************************************
    '*** 1 --- ADD
    '*** 2 --- SUBTRACT
    '*** 3 --- MULTIPLY
    '*** 4 --- DIVIDE
    '*** 5 --- PERCENTAGE
    '*** 6 --- X to Power n
    '*** 7 --- 1 Over x
    '*** 8 --- SQUARED
    '*** 9 --- SQUARE ROOT
    '*** 10 -- CUBED
    '******************************************************************
    '******DECLARE VALUES *********************************************
    '******************************************************************
    Dim num As Double = 0
    Dim num2 As Double = 0
    Dim KeyPressButt As Integer = 0
    Dim OpCode As Integer = 0
    Dim OpCode2 As Integer = 0
    Dim OpKey As Boolean = False
    Dim DP As Boolean = False
    '******************************************************************
    '******DECLARE STRINGS ********************************************
    '******************************************************************
    Dim numstr As String = "0"
    Dim numstr2 As String = "0"
    Dim total As String = "0"
    '******************************************************************
    '**** FUNCTIONS ***************************************************
    '******************************************************************
    Function CalcOp(OpCode) '*** Calculate When Operation Button Pushed
        OpKey = True
        numstr = TextBox1.Text
        If numstr = "" Then
            ResetVariables()
            MsgBox("Please input number first")
        Else
            num = Convert.ToDouble(numstr)
            total = num + num2
            '******************************************************************
            '***** PROCESS MATHS **********************************************
            ProcMaths(OpCode) '*** Perform Calculation ************************
            '******************************************************************
            TextBox1.Text = total
            'DeBugWindow()
            OpCode2 = OpCode
            DP = False
            num2 = total
            numstr2 = numstr
        End If
    End Function
    Function DeBugWindow()
        'DeBug.TextBox2.Text = num
        'DeBug.TextBox3.Text = num2
        'DeBug.TextBox4.Text = numstr
        'DeBug.TextBox5.Text = numstr2
        'DeBug.TextBox6.Text = total
        'DeBug.TextBox7.Text = OpCode
        'DeBug.TextBox8.Text = OpCode2
    End Function
    Function ProcMaths(OpCode) '***  Perform Calulation *****
        If OpCode2 = 1 Then
            total = Convert.ToString(num2 + num)
            Return total
        End If
        If OpCode2 = 2 Then
            total = Convert.ToString(num2 - num)
            Return total
        End If
        If OpCode2 = 3 Then
            total = Convert.ToString(num2 * num)
            Return total
        End If
        If OpCode2 = 4 Then
            total = Convert.ToString(num2 / num)
            Return total
        End If
        If OpCode2 = 5 Then
            total = (num2 / 100) * num
            Return total
        End If
        If OpCode2 = 6 Then
            total = Math.Pow(num2, num)
            Return total
        End If
        If OpCode2 = 7 Then
            total = Convert.ToString(1 / num)
            Return total
        End If
        If OpCode2 = 8 Then
            num = Convert.ToDouble(TextBox1.Text)
            'TextBox1.Text = (num * num)
            total = Convert.ToString(num2 * num)
            Return total
        End If
        If OpCode2 = 9 Then
            '**** SQUARE ROOT **********************************************
            total = Convert.ToString(Math.Sqrt(TextBox1.Text))
            Return total
        End If
        If OpCode2 = 10 Then
            total = Convert.ToString(num * num * num)
            Return total
        End If
    End Function
    Function InputNumber(KeyPressButt) '*** Process Number Button ****
        If OpKey = True Then
            TextBox1.Text = ""
            OpKey = False
        End If
        TextBox1.AppendText(Convert.ToString(KeyPressButt))
    End Function
    Function ResetVariables()
        num = 0
        num2 = 0
        KeyPressButt = 0
        OpKey = False
        DP = False
        OpCode = 0
        OpCode2 = 0
        '****** Declare Strings
        numstr = "0"
        numstr2 = "0"
        total = "0"
        TextBox1.Text = ""
        DeBug.TextBox2.Text = ""
        DeBug.TextBox3.Text = ""
        DeBug.TextBox4.Text = ""
        DeBug.TextBox5.Text = ""
        DeBug.TextBox6.Text = ""
        DeBug.TextBox7.Text = ""
        DeBug.TextBox8.Text = ""
    End Function
    '******************************************************************
    '**** OP BUTTONS **************************************************
    '******************************************************************
    Private Sub ADDBUTTON_Click(sender As Object, e As EventArgs) Handles Button3.Click
        CalcOp(1)
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        CalcOp(2)
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click
        CalcOp(3)
    End Sub
    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click
        CalcOp(4)
    End Sub
    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click
        CalcOp(5)
    End Sub
    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click
        CalcOp(6)
    End Sub
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click
        CalcOp(7)
    End Sub
    Private Sub Button22_Click(sender As Object, e As EventArgs) Handles Button22.Click
        CalcOp(8)
    End Sub
    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click
        CalcOp(9)
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click
        CalcOp(10)
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        '**** EQUAL BUTTON **********************************************
        CalcOp(OpCode2)
        num = 0
        num2 = 0
        numstr = 0
        numstr2 = 0
        OpCode2 = 0
        DeBugWindow()
    End Sub
    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click
        '**** CLEAR BUTTON **********************************************
        ResetVariables()
    End Sub
    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        '**** DECIMAL POINT BUTTON ************************************
        If DP = False Then
            TextBox1.AppendText(".")
            DP = True
        End If
    End Sub
    Private Sub Button16_Click_1(sender As Object, e As EventArgs) Handles Button16.Click
        '**** BACKSPACE **********************************************
        If TextBox1.Text.Length <> 0 Then
            numstr = TextBox1.Text
            TextBox1.Text = (TextBox1.Text.Substring(0, TextBox1.Text.Length - 1))
        End If
    End Sub
    '******************************************************************
    '**** NUMBER BUTTONS **********************************************
    '******************************************************************
    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        InputNumber(0)
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        InputNumber(1)
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        InputNumber(2)
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        InputNumber(3)
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        InputNumber(4)
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        InputNumber(5)
    End Sub
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        InputNumber(6)
    End Sub
    Private Sub num1button_Click(sender As Object, e As EventArgs) Handles Button1.Click
        InputNumber(7)
    End Sub
    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        InputNumber(8)
    End Sub
    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        InputNumber(9)
    End Sub
    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click
        '**** PI **********************************************
        TextBox1.Text = Math.PI
        OpKey = True
    End Sub
    Private Sub Button16_Click(sender As Object, e As EventArgs)
        '****** TEST BUTTON **********************************************
        OpKey = True
        numstr = TextBox1.Text
        num = Convert.ToDouble(numstr)
        total = num + num2
        TextBox1.Text = total
        DeBugWindow()
        num2 = total
        numstr2 = numstr
    End Sub
    Private Sub Button26_Click(sender As Object, e As EventArgs) Handles Button26.Click
        '*** Copy to clipboard
        Clipboard.SetText(TextBox1.Text)
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ResetVariables()
        'DeBug.Show()
    End Sub
    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click
        '*** Copy from clipboard
        TextBox1.Text = Clipboard.GetText
    End Sub
    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click
        '*** SIGN CHANGE **********************************************
        TextBox1.Text = TextBox1.Text * (-1)
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click
        '** Clear TextBox
        TextBox1.Text = ""
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItem1.Click

    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutToolStripMenuItem.Click
        frmAboutCalculator.Show()
    End Sub
End Class
