Partial Class SapFiPlAddIn
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(SapFiPlAddIn))
        Me.SapFiPl = Me.Factory.CreateRibbonTab
        Me.SAPNewGLplanning = Me.Factory.CreateRibbonGroup
        Me.ButtonNewGLplanningCheck = Me.Factory.CreateRibbonButton
        Me.ButtonNewGLplanningPost = Me.Factory.CreateRibbonButton
        Me.SAPFiPlLogon = Me.Factory.CreateRibbonGroup
        Me.ButtonLogon = Me.Factory.CreateRibbonButton
        Me.ButtonLogoff = Me.Factory.CreateRibbonButton
        Me.SapFiPl.SuspendLayout()
        Me.SAPNewGLplanning.SuspendLayout()
        Me.SAPFiPlLogon.SuspendLayout()
        Me.SuspendLayout()
        '
        'SapFiPl
        '
        Me.SapFiPl.Groups.Add(Me.SAPNewGLplanning)
        Me.SapFiPl.Groups.Add(Me.SAPFiPlLogon)
        Me.SapFiPl.Label = "SAP FI-Pl"
        Me.SapFiPl.Name = "SapFiPl"
        '
        'SAPNewGLplanning
        '
        Me.SAPNewGLplanning.Items.Add(Me.ButtonNewGLplanningCheck)
        Me.SAPNewGLplanning.Items.Add(Me.ButtonNewGLplanningPost)
        Me.SAPNewGLplanning.Label = "New-GL Planning"
        Me.SAPNewGLplanning.Name = "SAPNewGLplanning"
        '
        'ButtonNewGLplanningCheck
        '
        Me.ButtonNewGLplanningCheck.Image = CType(resources.GetObject("ButtonNewGLplanningCheck.Image"), System.Drawing.Image)
        Me.ButtonNewGLplanningCheck.Label = "NewGLplanning Check"
        Me.ButtonNewGLplanningCheck.Name = "ButtonNewGLplanningCheck"
        Me.ButtonNewGLplanningCheck.ShowImage = True
        '
        'ButtonNewGLplanningPost
        '
        Me.ButtonNewGLplanningPost.Image = CType(resources.GetObject("ButtonNewGLplanningPost.Image"), System.Drawing.Image)
        Me.ButtonNewGLplanningPost.Label = "NewGLplanning Post"
        Me.ButtonNewGLplanningPost.Name = "ButtonNewGLplanningPost"
        Me.ButtonNewGLplanningPost.ShowImage = True
        '
        'SAPFiPlLogon
        '
        Me.SAPFiPlLogon.Items.Add(Me.ButtonLogon)
        Me.SAPFiPlLogon.Items.Add(Me.ButtonLogoff)
        Me.SAPFiPlLogon.Label = "Logon"
        Me.SAPFiPlLogon.Name = "SAPFiPlLogon"
        '
        'ButtonLogon
        '
        Me.ButtonLogon.Image = CType(resources.GetObject("ButtonLogon.Image"), System.Drawing.Image)
        Me.ButtonLogon.Label = "SAP Logon"
        Me.ButtonLogon.Name = "ButtonLogon"
        Me.ButtonLogon.ShowImage = True
        '
        'ButtonLogoff
        '
        Me.ButtonLogoff.Image = CType(resources.GetObject("ButtonLogoff.Image"), System.Drawing.Image)
        Me.ButtonLogoff.Label = "SAP Logoff"
        Me.ButtonLogoff.Name = "ButtonLogoff"
        Me.ButtonLogoff.ShowImage = True
        '
        'SapFiPlAddIn
        '
        Me.Name = "SapFiPlAddIn"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.SapFiPl)
        Me.SapFiPl.ResumeLayout(False)
        Me.SapFiPl.PerformLayout()
        Me.SAPNewGLplanning.ResumeLayout(False)
        Me.SAPNewGLplanning.PerformLayout()
        Me.SAPFiPlLogon.ResumeLayout(False)
        Me.SAPFiPlLogon.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents SapFiPl As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents SAPNewGLplanning As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents SAPFiPlLogon As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ButtonLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonLogoff As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonNewGLplanningCheck As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ButtonNewGLplanningPost As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property SapFiPlAddIn() As SapFiPlAddIn
        Get
            Return Me.GetRibbon(Of SapFiPlAddIn)()
        End Get
    End Property
End Class
