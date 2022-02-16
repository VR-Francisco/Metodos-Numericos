Public Class primero
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label26_Click(sender As Object, e As EventArgs)

    End Sub



    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Parcial2.Show()
        Me.Hide()
    End Sub

    Private Sub PictureBox2_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs)
        Me.Hide()
        Main.Show()
    End Sub

    Private Sub picMaximize_Click(sender As Object, e As EventArgs) Handles picMaximize.Click
        If Me.WindowState = WindowState.Normal Then
            Me.WindowState = WindowState.Maximized
        ElseIf Me.WindowState = WindowState.Maximized Then
            Me.WindowState = WindowState.Normal
        End If
    End Sub

    Private Sub picMinimize_Click(sender As Object, e As EventArgs) Handles picMinimize.Click
        Me.WindowState = WindowState.Minimized
    End Sub

    Private Sub picClose_Click(sender As Object, e As EventArgs) Handles picClose.Click
        End
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        pnlposition.Height = Button4.Height
        pnlposition.Top = Button4.Top
        Me.Hide()
        Main.Show()
        Inicio.Visible = True
        AcercaDe.Visible = False
    End Sub

    Private Sub Button1_Click_1(sender As Object, e As EventArgs) Handles Button1.Click
        pnlposition.Height = Button1.Height
        pnlposition.Top = Button1.Top
        Parcial2.Show()
        Me.Hide()
        Inicio.Visible = True
        AcercaDe.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click


    End Sub

    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        pnlposition.Height = Button3.Height
        pnlposition.Top = Button3.Top
        AcercaDe.Visible = True
        Inicio.Visible = False
        Panel2.Visible = False
    End Sub

    Private Sub Home_Click(sender As Object, e As EventArgs) Handles Home.Click
        pnlposition.Height = Home.Height
        pnlposition.Top = Home.Top
        AcercaDe.Visible = False
        Inicio.Visible = True
        Panel2.Visible = False
    End Sub

    Private Sub AcercaDe_Paint(sender As Object, e As PaintEventArgs) Handles AcercaDe.Paint

    End Sub
End Class
