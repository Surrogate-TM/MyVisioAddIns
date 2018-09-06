Public Class CFP
    Private Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOK.Click
        MsgBox("Вы должны выбрать одну фигуру")
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        Me.Close()
    End Sub

    Private Sub CFP_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
End Class