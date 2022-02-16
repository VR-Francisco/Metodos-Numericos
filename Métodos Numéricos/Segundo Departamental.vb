Public Class Parcial2
    Dim x(60), Y(60), z(60), ec, ex(60), ey(60), ez(60), s, sk(60), deltay(60), fact, prod, prod1, suma, yres, xl(60), yl(60), xl2(60), yl2(60), yxl, yxl2, vxl, vxl2, nm, a, b, xm(60), ym(60), sumaX, sumaY, sumaXY, sumaXX, xMedia, yMedia, vxmin As Single

    Private Sub Label48_Click(sender As Object, e As EventArgs) Handles lby.Click

    End Sub

    Private Sub btmin3_Click(sender As Object, e As EventArgs) Handles btmin3.Click
        g = CreateGraphics()
        For i = 1 To nm Step 1
            graf.Series(0).Points.AddXY(xm(i), ym(i))
        Next

        graf.Series(1).Points.AddXY(xmin, ymin)
        graf.Series(1).Points.AddXY(xmax, ymax)
    End Sub

    Dim i = 0, redon, c, facti, r, k, nl, nl2, il, il2, ibl, ibl2, cm As Integer
    Dim g As Graphics
    Dim xmin, ymin, xmax, ymax As Single

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click
        pnlposition.Height = Button14.Height
        pnlposition.Top = Button14.Top
        Jacobi.Visible = False
        Gauss.Visible = False
        Interpolación.Visible = False
        Inicio.Visible = False
        lagrange.Visible = False
        Lagrange2.Visible = False
        Mínimos.Visible = True
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles btmin1.Click
        nm = minCua1.Text
        cm = minCua2.Text
        vxmin = vxm.Text

        redon = cm + 2

        sumaXY = 0
        sumaXX = 0
        sumaX = 0
        sumaY = 0

        For i = 1 To nm Step 1
            xm(i) = InputBox("x(" & i & ")=")
            ym(i) = InputBox("y(" & i & ")=")
            Salmin.Rows.Add(i, xm(i), ym(i))

            'Suma x
            sumaX = sumaX + xm(i)

            'Suma y
            sumaY = sumaY + ym(i)

            'Suma xy
            sumaXY = sumaXY + xm(i) * ym(i)

            'Suma x cuadrada
            sumaXX = sumaXX + xm(i) * xm(i)
        Next

        xMedia = sumaX / nm
        yMedia = sumaY / nm

        b = (sumaXY - nm * xMedia * yMedia) / (sumaXX - nm * xMedia ^ 2)
        a = yMedia - b * xMedia

        Label43.Text = a + b * vxmin

        xmin = xm(1)
        ymin = a + b * xmin
        xmax = xm(nm)
        ymax = a + b * xmax

        lba.Text = Math.Round(a, redon)
        lbb.Text = Math.Round(b, redon)


    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        SalidaLag.Rows.Clear()
        txp.Clear()
        txi.Clear()
        reslag.Clear()
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click
        Salidal2.Rows.Clear()
        l2par.Clear()
        l2vx.Clear()
        l2res.Clear()
    End Sub

    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        nl2 = l2par.Text
        vxl2 = l2vx.Text


        For il2 = 0 To nl2 - 1 Step 1
            xl2(il2) = InputBox("X(" & il2 & ")=")
            yl2(il2) = InputBox("Y(" & il2 & ")=")
            Salidal2.Rows.Add(il2, xl2(il2), yl2(il2))
        Next

        il2 = 0
        Do While vxl2 > xl2(il2)
            il2 = il2 + 1
        Loop
        ibl2 = il2 - 1

        yxl2 = (((vxl2 - xl2(ibl2 + 1)) * (vxl2 - xl2(ibl2 + 2))) / ((xl2(ibl2) - xl2(ibl2 + 1)) * (xl2(ibl2) - xl2(ibl2 + 2)))) * yl2(ibl2) + (((vxl2 - xl2(ibl2)) * (vxl2 - xl2(ibl2 + 2))) / ((xl2(ibl2 + 1) - xl2(ibl2)) * (xl2(ibl2 + 1) - xl2(ibl2 + 2)))) * yl2(ibl2 + 1) + (((vxl2 - xl2(ibl2)) * (vxl2 - xl2(ibl2 + 1))) / ((xl2(ibl2 + 2) - xl2(ibl2)) * (xl2(ibl2 + 2) - xl2(ibl2 + 1)))) * yl2(ibl2 + 2)
        l2res.Text = yxl2
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        pnlposition.Height = Button11.Height
        pnlposition.Top = Button11.Top
        Jacobi.Visible = False
        Gauss.Visible = False
        Interpolación.Visible = False
        Inicio.Visible = False
        lagrange.Visible = False
        Lagrange2.Visible = True
        Mínimos.Visible = False
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        pnlposition.Height = Button9.Height
        pnlposition.Top = Button9.Top
        Jacobi.Visible = False
        Gauss.Visible = False
        Interpolación.Visible = False
        Inicio.Visible = False
        lagrange.Visible = True
        Lagrange2.Visible = False
        Mínimos.Visible = False
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        nl = txp.Text
        vxl = txi.Text


        For il = 0 To nl - 1 Step 1
            xl(il) = InputBox("X(" & il & ")=")
            yl(il) = InputBox("Y(" & il & ")=")
            SalidaLag.Rows.Add(il, xl(il), yl(il))
        Next

        il = 0
        Do While vxl > xl(il)
            il = il + 1
        Loop
        ibl = il - 1

        yxl = (((vxl - xl(ibl + 1)) / (xl(ibl) - xl(ibl + 1))) * yl(ibl)) + (((vxl - xl(ibl)) / (xl(ibl + 1) - xl(ibl))) * yl(ibl + 1))
        reslag.Text = yxl
    End Sub

    Dim n, vx, ib, m As Integer
    Dim j As Double
    Private Sub PictureBox5_Click(sender As Object, e As EventArgs)
        primero.Show()
        Me.Hide()
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        i = 0
        c = TextBox13.Text
        redon = c + 2
        ec = 0.5 * 10 ^ (-c)
        ex(i) = 1
        ey(i) = 1
        ez(i) = 1

        DataGridView1.Rows.Add(i, Math.Round(x(i), redon), Math.Round(Y(i), redon),
            Math.Round(z(i), redon), "-----", "-----", "-----")

        Do While ex(i) > ec Or ey(i) > ec Or ez(i) > ec
            i = i + 1

            x(i) = (b11.Text - cont2.Text * Y(i - 1) - cont3.Text * z(i - 1)) / cont1.Text
            Y(i) = (b22.Text - cont11.Text * x(i) - cont31.Text * z(i - 1)) / cont21.Text
            z(i) = (b33.Text - cont111.Text * x(i) - cont222.Text * Y(i)) / cont333.Text

            ex(i) = Math.Abs((x(i) - x(i - 1)) / x(i))
            ey(i) = Math.Abs((Y(i) - Y(i - 1)) / Y(i))
            ez(i) = Math.Abs((z(i) - z(i - 1)) / z(i))
            DataGridView1.Rows.Add(i, Math.Round(x(i), redon), Math.Round(Y(i), redon), Math.Round(z(i), redon), Math.Round(ex(i), redon), Math.Round(ey(i), redon), Math.Round(ez(i), redon))

        Loop
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        DataGridView1.Rows.Clear()
        cont1.Clear()
        cont11.Clear()
        cont111.Clear()
        cont2.Clear()
        cont21.Clear()
        cont222.Clear()
        cont3.Clear()
        cont31.Clear()
        cont333.Clear()
        b11.Clear()
        b22.Clear()
        b33.Clear()
        TextBox13.Clear()
    End Sub

    Private Sub Limpiar_Click(sender As Object, e As EventArgs) Handles Limpiar.Click
        i1.Clear()
        i2.Clear()
        i3.Clear()
        i4.Clear()
        i5.Clear()
        Sinter.Rows.Clear()
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Interpolación.Paint

    End Sub

    Private Sub Label21_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        pnlposition.Height = Button7.Height
        pnlposition.Top = Button7.Top
        Jacobi.Visible = False
        Gauss.Visible = False
        Interpolación.Visible = True
        Inicio.Visible = False
        lagrange.Visible = False
        Lagrange2.Visible = False
        Mínimos.Visible = False

    End Sub

    Private Sub Calcular_Click(sender As Object, e As EventArgs) Handles Calcular.Click
        vx = i2.Text
        n = i1.Text

        For i = 0 To n - 1 Step 1
            x(i) = InputBox("X(" & i & ")=")
            Y(i) = InputBox("Y(" & i & ")=")
            Sinter.Rows.Add(i, x(i), Y(i))
        Next

        i = 0
        Do While vx > x(i)
            i = i + 1
        Loop

        ib = i - 1
        m = n - (ib + 1)
        s = (vx - x(ib)) / (x(ib + 1) - x(ib))

        'Calculo de los coeficientes de s
        'fact = 1
        'sk(0) = 1
        'sk(1) = s
        'For i = 2 To m Step 1
        'For facti = 1 To i Step 1
        'fact *= facti
        'Next
        'sk(i) = (s * (s - 1)) / fact
        'Next

        For k = 0 To m Step 1
            prod = 1
            For r = 1 To k Step 1
                prod = prod * ((s - (r - 1)) / r)
            Next
            sk(k) = prod
            cofiBi.Rows.Add("", "", "", sk(k))
        Next

        For k = 0 To m Step 1
            suma = 0
            For j = 0 To k
                prod1 = 1
                For r = 1 To j
                    prod1 = prod1 * (k - (r - 1)) / r
                Next
                suma = suma + ((-1) ^ j) * prod1 * (Y(ib + k - j))
            Next
            deltay(k) = suma
            difDelta.Rows.Add("", "", "", "", Math.Round(Convert.ToDouble(deltay(k)), 4))
        Next

        yres = 0
        For k = 0 To m Step 1
            yres = yres + sk(k) * deltay(k)
        Next

        TextBox1.Text = Math.Round(Convert.ToDouble(yres), 4)
        i4.Text = m
        i5.Text = s
        i3.Text = ib

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        pnlposition.Height = Button1.Height
        pnlposition.Top = Button1.Top
        Jacobi.Visible = True
        Gauss.Visible = False
        Interpolación.Visible = False
        Inicio.Visible = False
        lagrange.Visible = False
        Lagrange2.Visible = False
        Mínimos.Visible = False
    End Sub

    Private Sub picClose_Click(sender As Object, e As EventArgs) Handles picClose.Click
        End
    End Sub

    Private Sub Barra_Paint(sender As Object, e As PaintEventArgs) Handles BarraIzq.Paint

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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        pnlposition.Height = Button2.Height
        pnlposition.Top = Button2.Top
        Jacobi.Visible = False
        Gauss.Visible = True
        Interpolación.Visible = False
        Inicio.Visible = False
        lagrange.Visible = False
        Lagrange2.Visible = False
        Mínimos.Visible = False
    End Sub

    Private Sub Home_Click(sender As Object, e As EventArgs) Handles Home.Click
        pnlposition.Height = Home.Height
        pnlposition.Top = Home.Top
        Me.Hide()
        primero.Show()
        Jacobi.Visible = False
        Gauss.Visible = False
        Interpolación.Visible = False
        Inicio.Visible = True
        lagrange.Visible = False
        Lagrange2.Visible = False
        Mínimos.Visible = False
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        salida.Rows.Clear()
        a11.Clear()
        a12.Clear()
        a13.Clear()
        a21.Clear()
        a22.Clear()
        a23.Clear()
        a31.Clear()
        a32.Clear()
        a33.Clear()
        b1.Clear()
        b2.Clear()
        b3.Clear()
        tc.Clear()
        lbx.Text = "Sin calcular"
        lby.Text = "Sin calcular"
        lbz.Text = "Sin calcular"
        lbredon.Text = " "
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        i = 0
        c = tc.Text
        redon = c + 2
        ec = 0.5 * 10 ^ (-c)
        ex(i) = 1
        ey(i) = 1
        ez(i) = 1

        salida.Rows.Add(i, Math.Round(x(i), redon), Math.Round(Y(i), redon),
            Math.Round(z(i), redon), "-----", "-----", "-----")

        Do While ex(i) > ec Or ey(i) > ec Or ez(i) > ec
            i = i + 1

            x(i) = (b1.Text - a12.Text * Y(i - 1) - a13.Text * z(i - 1)) / a11.Text
            Y(i) = (b2.Text - a21.Text * x(i - 1) - a23.Text * z(i - 1)) / a22.Text
            z(i) = (b3.Text - a31.Text * x(i - 1) - a32.Text * Y(i - 1)) / a33.Text

            ex(i) = Math.Abs((x(i) - x(i - 1)) / x(i))
            ey(i) = Math.Abs((Y(i) - Y(i - 1)) / Y(i))
            ez(i) = Math.Abs((z(i) - z(i - 1)) / z(i))
            salida.Rows.Add(i, Math.Round(x(i), redon), Math.Round(Y(i), redon), Math.Round(z(i), redon), Math.Round(ex(i), redon), Math.Round(ey(i), redon), Math.Round(ez(i), redon))
        Loop

        lbredon.Text = ec

        lbx.Text = x(i)
        lby.Text = Y(i)
        lbz.Text = z(i)

    End Sub
End Class