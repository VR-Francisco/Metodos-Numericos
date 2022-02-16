Imports info.lundin.math

Public Class Main
    Dim a, a1, a2 As Single
    Dim b, b1, b2 As Single
    Dim c, c1, c2, cifr As Integer
    Dim x(50), x1(100), m, x2(100), xr(100) As Single
    Dim err(50), err1(100), err2(100), errrz(100) As Single
    Dim ec, ec1, ec2, ecr As Single
    Dim i, i1, i2, irz, j, impr, fin, redonrz As Integer

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        pnlposition.Height = Button3.Height
        pnlposition.Top = Button3.Top
        Bisección.Visible = False
        ReglaFalsa.Visible = False
        NewtonRaphson.Visible = True
        Inicio.Visible = False
        Impares.Visible = False
        raizdedos.Visible = False
    End Sub

    Private Sub picClose_Click(sender As Object, e As EventArgs) Handles picClose.Click
        End
    End Sub

    Private Sub picMinimize_Click(sender As Object, e As EventArgs) Handles picMinimize.Click
        Me.WindowState = WindowState.Minimized
    End Sub

    Private Sub picMaximize_Click(sender As Object, e As EventArgs) Handles picMaximize.Click
        If Me.WindowState = WindowState.Normal Then
            Me.WindowState = WindowState.Maximized
        ElseIf Me.WindowState = WindowState.Maximized Then
            Me.WindowState = WindowState.Normal
        End If
    End Sub

    Private Sub btclear_Click(sender As Object, e As EventArgs) Handles btclear.Click
        raiztb1.Clear()
        raiztb2.Clear()
        gridraiz.Rows.Clear()
        irz = 0
    End Sub

    Private Sub Label25_Click(sender As Object, e As EventArgs) Handles Label25.Click

    End Sub

    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)
        primero.Show()
        Me.Hide()
    End Sub

    Private Sub Home_Click(sender As Object, e As EventArgs) Handles Home.Click
        pnlposition.Height = Home.Height
        pnlposition.Top = Home.Top
        Me.Hide()
        primero.Show()
        Bisección.Visible = False
        ReglaFalsa.Visible = False
        NewtonRaphson.Visible = False
        Inicio.Visible = True
        Impares.Visible = False
        raizdedos.Visible = False
    End Sub

    Private Sub impar_Click(sender As Object, e As EventArgs) Handles impar.Click
        pnlposition.Height = impar.Height
        pnlposition.Top = impar.Top
        Bisección.Visible = False
        ReglaFalsa.Visible = False
        NewtonRaphson.Visible = False
        Inicio.Visible = False
        Impares.Visible = True
        raizdedos.Visible = False
    End Sub

    Private Sub raizdos_Click(sender As Object, e As EventArgs) Handles raizdos.Click
        pnlposition.Height = raizdos.Height
        pnlposition.Top = raizdos.Top
        Bisección.Visible = False
        ReglaFalsa.Visible = False
        NewtonRaphson.Visible = False
        Inicio.Visible = False
        Impares.Visible = False
        raizdedos.Visible = True
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles btraiz.Click
        xr(irz) = raiztb1.Text
        cifr = raiztb2.Text
        ecr = 0.5 * 10 ^ (-cifr)
        errrz(irz) = 1
        redonrz = cifr + 2
        Do While errrz(irz) > ecr
            irz = irz + 1
            xr(irz) = 0.5 * (xr(irz - 1) + 2 / xr(irz - 1))
            errrz(irz) = Math.Abs((xr(irz) - xr(irz - 1)) / xr(irz))
            gridraiz.Rows.Add(irz, Math.Round(xr(irz), redonrz), Math.Round(errrz(irz), redonrz))
        Loop
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        imp.Clear()
        salimpar.Rows.Clear()
    End Sub

    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        fin = imp.Text
        For k = 1 To fin
            impr = 2 * k - 1
            salimpar.Rows.Add(k, impr)
        Next
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        pnlposition.Height = Button2.Height
        pnlposition.Top = Button2.Top
        Bisección.Visible = False
        ReglaFalsa.Visible = True
        NewtonRaphson.Visible = False
        Inicio.Visible = False
        Impares.Visible = False
        raizdedos.Visible = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        pnlposition.Height = Button1.Height
        pnlposition.Top = Button1.Top
        Bisección.Visible = True
        ReglaFalsa.Visible = False
        NewtonRaphson.Visible = False
        Inicio.Visible = False
        Impares.Visible = False
        raizdedos.Visible = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        c2 = tc2.Text
        a2 = ta2.Text
        b2 = tb2.Text
        j = 1
        m = (a2 + b2) / 2
        redon2 = c2 + 2
        ec2 = 0.5 * 10 ^ (-c2)
        x2(i2) = m - (f2(m) / d(m))
        err2(i2) = 1
        Salida2.Rows.Add(i2, m, "-------")
        Salida2.Rows.Add(j, Math.Round(x2(i2), redon2), Math.Round((x2(i2) - m) / x2(i2), redon2))
        Do While err2(i2) > ec2
            x2(i2 + 1) = x2(i2) - (f2(x2(i2)) / d(x2(i2)))
            err2(i2 + 1) = Math.Abs((x2(i2 + 1) - x2(i2)) / x2(i2 + 1))
            Salida2.Rows.Add(j + 1, Math.Round(x2(i2 + 1), redon2), Math.Round(err2(i2 + 1), redon2))
            j = j + 1
            i2 = i2 + 1
        Loop

        Salida2.Rows.Add("La raíz es: ", Math.Round(x2(i2), redon2))
        Raiz2.Visible = True
        res2.Visible = True
        res2.Text = Convert.ToString(Math.Round(x2(i2), redon2))
    End Sub

    Dim redon, redon1, redon2 As Integer

    Function f(x As Single) As Single
        Dim parser As ExpressionParser
        parser = New ExpressionParser
        parser.Values.Clear()
        parser.Values.Add("x", x)
        Return parser.Parse(tf.Text)
    End Function

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        tf2.Clear()
        tfd.Clear()
        ta2.Clear()
        tb2.Clear()
        tc2.Clear()
        Salida2.Rows.Clear()
        Raiz2.Visible = False
        res2.Visible = False
    End Sub

    Function f1(x1 As Single) As Single
        Dim parser As ExpressionParser
        parser = New ExpressionParser
        parser.Values.Clear()
        parser.Values.Add("x", x1)
        Return parser.Parse(tf1.Text)
    End Function

    Function f2(x2 As Single) As Single
        Dim parser As ExpressionParser
        parser = New ExpressionParser
        parser.Values.Clear()
        parser.Values.Add("x", x2)
        Return parser.Parse(tf2.Text)
    End Function

    Function d(y As Single) As Single

        Dim parser As ExpressionParser
        parser = New ExpressionParser
        parser.Values.Clear()
        parser.Values.Add("x", y)
        Return parser.Parse(tfd.Text)

    End Function

    Private Sub Panel6_Paint(sender As Object, e As PaintEventArgs) Handles pnlposition.Paint

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        End
    End Sub

    Private Sub PictureBox1_Click(sender As Object, e As EventArgs) Handles PictureBox1.Click

    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        a = ta.Text
        b = tb.Text
        c = tc.Text
        ec = 0.5 * 10 ^ (-c)
        redon = c + 2
        x(i) = (a + b) / 2
        err(i) = 1
        i = 0

        If f(x(i)) = 0 Then
            err(i) = 0
        End If
        Salida.Rows.Add(i, Math.Round(a, redon), Math.Round(x(i), redon),
                        Math.Round(b, redon), Math.Round(f(a), redon),
                        Math.Round(f(x(i)), redon), Math.Round(f(b), redon),
                        "-------")

        Do While err(i) > ec
            If f(a) * f(x(i)) < 0 Then
                b = x(i)
            Else
                a = x(i)
            End If
            i = i + 1
            x(i) = (a + b) / 2
            err(i) = Math.Abs((x(i) - x(i - 1)) / x(i))

            Salida.Rows.Add(i, Math.Round(a, redon), Math.Round(x(i), redon),
           Math.Round(b, redon),
            Math.Round(f(a), redon), Math.Round(f(x(i)), redon), Math.Round(f(b), redon),
           Math.Round(err(i), redon))
        Loop
        Salida.Rows.Add("La raíz es: ", Math.Round(x(i), redon))
        raiz.Visible = True
        Res.Visible = True
        Res.Text = Convert.ToString(Math.Round(x(i), redon))
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        tf.Clear()
        ta.Clear()
        tb.Clear()
        tc.Clear()
        Salida.Rows.Clear()
        raiz.Visible = False
        Res.Visible = False
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        a1 = ta1.Text
        b1 = tb1.Text
        c1 = tc1.Text
        redon1 = c1 + 2
        ec1 = 0.5 * 10 ^ (-c1)
        i1 = 0
        err1(i1) = 1

        x1(i1) = (a1 * f1(b1) - b1 * f1(a1)) / (f1(b1) - f1(a1))

        If f1(x1(i1)) = 0 Then
            err1(i1) = 0
        End If
        Salida1.Rows.Add(i1, Math.Round(a1, redon1), Math.Round(x1(i1), redon1),
                        Math.Round(b1, redon1), Math.Round(f1(a1), redon1),
                        Math.Round(f1(x1(i1)), redon1), Math.Round(f1(b1), redon1),
                        "-------")

        Do While err1(i1) > ec1
            If f1(a1) * f1(x1(i1)) < 0 Then
                b1 = x1(i1)
            Else
                a1 = x1(i1)
            End If
            i1 = i1 + 1

            x1(i1) = (a1 * f1(b1) - b1 * f1(a1)) / (f1(b1) - f1(a1))
            err1(i1) = Math.Abs((x1(i1) - x1(i1 - 1)) / x1(i1))
            Salida1.Rows.Add(i1, Math.Round(a1, redon1), Math.Round(x1(i1), redon1),
           Math.Round(b1, redon1),
            Math.Round(f1(a1), redon1), Math.Round(f1(x1(i1)), redon1), Math.Round(f1(b1), redon1),
           Math.Round(err1(i1), redon1))
        Loop
        Salida1.Rows.Add("La raíz es: ", Math.Round(x1(i1), redon1))
        Raiz1.Visible = True
        res1.Visible = True
        res1.Text = Convert.ToString(Math.Round(x1(i1), redon1))
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        tf1.Clear()
        ta1.Clear()
        tb1.Clear()
        tc1.Clear()
        Salida1.Rows.Clear()
        Raiz1.Visible = False
        res1.Visible = False
    End Sub
End Class
