Imports System.Drawing
Imports System.Windows.Forms

Module xxx

#Region "LIST OF VARIABLES AND CONSTANTS"
    Friend frmNewTable As System.Windows.Forms.Form
    Friend frmCFP As System.Windows.Forms.Form

#End Region

#Region "Load Sub"

    Sub Load_CFP()
        If frmCFP Is Nothing Then
            frmCFP = New CFP
            frmCFP.Show()
        End If
    End Sub

#End Region

#Region "Friend Sub"


#End Region

End Module
