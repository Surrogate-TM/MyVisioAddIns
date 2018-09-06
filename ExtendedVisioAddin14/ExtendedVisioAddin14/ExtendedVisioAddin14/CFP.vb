Imports System.Windows.Forms

Public Class CFP
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Private Sub BtnOk_Click(sender As Object, e As EventArgs) Handles BtnOK.Click
        MsgBox("OK")
        Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        'Me.Hide()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Dim cFP As CFP = Me
        'cFP.Close()
        Me.Close()
    End Sub

    Private Sub CFP_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub


    Private Sub CFP_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        xxx.frmCFP = Nothing
    End Sub
End Class