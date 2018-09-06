Imports System.Drawing
Imports System.Windows.Forms
Imports Office = Microsoft.Office.Core
Imports Visio = Microsoft.Office.Interop.Visio
Partial Public Class ThisAddIn
    Public Property AddinUI As AddinUI = new AddinUI()

    Protected Overrides Function CreateRibbonExtensibilityObject() As Office.IRibbonExtensibility
        Return _addinUI
    End Function

    ''' 
    ''' A simple command
    ''' 

    ''' 
    ''' Callback called by the UI manager when user clicks a button
    ''' Should do something meaninful wehn corresponding action is called.
    ''' 
    Public Sub OnCommand(commandId As String)
        Select Case commandId
            Case "BtnCFP"
                Load_CFP()
                Return
            Case "BtnPDF" : MsgBox("Export to PDF") : Return
            Case "BtnImp" : MsgBox("Import specification") : Return
            Case "BtnExp" : MsgBox("Export specification") : Return
            Case "BtnDel" : MsgBox("Delete specification") : Return
                'Case "Btn_newtable"
                '    Btn_newtable()
                '    Return

        End Select
    End Sub

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command shoudl be enabled in the user interface.
    ''' By default, all commands are enabled.
    ''' 
    Public Function IsCommandEnabled(commandId As String) As Boolean
        Select Case commandId
            Case "BtnCFP"
                ' make command1 always enabled
                Return True
            Case "BtnPDF"
                ' make command1 always enabled
                Return True
            Case "BtnImp"
                ' make command1 always enabled
                Return True
            Case "BtnExp"
                ' make command1 always enabled
                Return True
            Case "BtnDel"
                ' make command1 always enabled
                Return True
                'Case "Btn_newtable"
                '    ' make command2 enabled only if a window is opened
                '    Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing AndAlso Application.ActiveWindow.Selection.Count > 0
            Case Else
                Return True
        End Select
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Should return if corresponding command (button) is pressed or not (makes sense for toggle buttons)
    ''' 
    Public Function IsCommandChecked(command As String) As Boolean
        Return False
    End Function

    ''' 
    ''' Callback called by UI manager.
    ''' Returns a label associated with given command.
    ''' We assume for simplicity taht command labels are named simply named as [commandId]_Label (see resources)
    ''' 
    Public Function GetCommandLabel(command As String) As String
        Return My.Resources.ResourceManager.GetString(command & "_Label")
    End Function

    ''' 
    ''' Returns a icon associated with given command.
    ''' We assume for simplicity that icon ids are named after command commandId.
    ''' 
    Public Function GetCommandBitmap(command As String) As Bitmap
        Return DirectCast(My.Resources.ResourceManager.GetObject(command), Bitmap)
    End Function


    Sub UpdateUI()
        AddinUI.UpdateRibbon()
    End Sub

    Public Sub Application_SelectionChanged(window As Visio.Window)
        UpdateUI()
    End Sub

    Private Sub ThisAddIn_Startup() Handles Me.Startup
        AddHandler Application.SelectionChanged, AddressOf Application_SelectionChanged

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown
        RemoveHandler Application.SelectionChanged, AddressOf Application_SelectionChanged

    End Sub

End Class
