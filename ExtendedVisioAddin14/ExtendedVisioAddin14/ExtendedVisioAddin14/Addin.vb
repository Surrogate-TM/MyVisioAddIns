Imports System.Drawing
Imports System.Windows.Forms


Partial Public Class Addin
    Public Property Application As Microsoft.Office.Interop.Visio.Application
    'Dim Application As Microsoft.Office.Interop.Visio.Application
    Public Sub OnAction(commandId As String, commandTag As String)
        ' Очищать?
        Dim winObj As Microsoft.Office.Interop.Visio.Window = Application.ActiveWindow
        Dim docObj As Microsoft.Office.Interop.Visio.Document = Application.ActiveDocument
        Dim pagObj As Microsoft.Office.Interop.Visio.Page = Application.ActivePage
        Dim shpsObj As Microsoft.Office.Interop.Visio.Shapes = pagObj.Shapes
        ' Очищать?

        Select Case commandId
            'Case "btn_newtable" : Load_dlgNewTable() : Return
            Case "btn_CFP" : xxx.Load_CFP() : Return
        End Select

    End Sub

#Region "RIBBONFUNCTIONS"

    Public Function IsCommandAltEnabled(commandId As String) As Boolean
        Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
    End Function

    Public Function IsCommandEnabled(commandId As String) As Boolean
        Return Application IsNot Nothing AndAlso Application.ActiveWindow IsNot Nothing
        'If Application.ActiveWindow.Selection.Count > 0 AndAlso _
        '    Application.ActiveWindow.Selection(1).CellExistsU("User.TableName", 0) Then Return True
        'Return False
    End Function

    Sub Startup(app As Object)
        Application = DirectCast(app, Microsoft.Office.Interop.Visio.Application)
        System.Windows.Forms.Application.SetCompatibleTextRenderingDefault(False)
        AddHandler Application.DocumentCreated, AddressOf Application_DocumentListChanged
        AddHandler Application.DocumentOpened, AddressOf Application_DocumentListChanged
        AddHandler Application.BeforeDocumentClose, AddressOf Application_DocumentListChanged
    End Sub

    Private Sub Application_DocumentListChanged(ByVal doc As Microsoft.Office.Interop.Visio.Document)
        UpdateRibbon()
    End Sub

    Private Sub UpdateRibbon()
        Throw New NotImplementedException()
    End Sub

    Sub Application_ShapeAdded(ByVal Sh As Microsoft.Office.Interop.Visio.Shape)

    End Sub

#End Region

End Class