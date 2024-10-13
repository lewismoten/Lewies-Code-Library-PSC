Imports System.ComponentModel
Imports System.Configuration.Install
Imports Microsoft.Win32
Imports System.Reflection

' In order for this installer to work during your MSI setup files
' installation you must view the custom actions of your setup 
' project.  Add the project primary output under each node:
'	Install, Rollback, Commit, Uninstall

<RunInstaller(True)> Public Class ProjectInstaller
    Inherits System.Configuration.Install.Installer

#Region " Component Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Component Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Installer overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.  
    'Do not modify it using the code editor.
	Friend WithEvents ServiceProcessInstaller1 As System.ServiceProcess.ServiceProcessInstaller
	Friend WithEvents ServiceInstaller1 As System.ServiceProcess.ServiceInstaller
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.ServiceProcessInstaller1 = New System.ServiceProcess.ServiceProcessInstaller
		Me.ServiceInstaller1 = New System.ServiceProcess.ServiceInstaller
		'
		'ServiceProcessInstaller1
		'
		Me.ServiceProcessInstaller1.Account = System.ServiceProcess.ServiceAccount.LocalSystem
		Me.ServiceProcessInstaller1.Password = Nothing
		Me.ServiceProcessInstaller1.Username = Nothing
		'
		'ServiceInstaller1
		'
		Me.ServiceInstaller1.DisplayName = "Visual SourceSafe Monitor"
		Me.ServiceInstaller1.ServiceName = "Visual SourceSafe Journal Monitor"
		'
		'ProjectInstaller
		'
		Me.Installers.AddRange(New System.Configuration.Install.Installer() {Me.ServiceProcessInstaller1, Me.ServiceInstaller1})

	End Sub

#End Region

	Private Sub ServiceProcessInstaller1_AfterInstall(ByVal sender As System.Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles ServiceProcessInstaller1.AfterInstall

		MyBase.OnAfterInstall(e.SavedState)

		' Pull description from assembly attributes
		' See AssemblyInfo.vb
		Dim [Assembly] As System.Reflection.Assembly
		[Assembly] = System.Reflection.Assembly.GetExecutingAssembly
		Dim Description As String
		Dim Type As Type = GetType(System.Reflection.AssemblyDescriptionAttribute)
		Description = [Assembly].GetCustomAttributes(Type, True)(0).Description

		AddDescription(Me.ServiceInstaller1.ServiceName, Description)

	End Sub

	Private Sub AddDescription(ByVal serviceName As String, ByVal description As String)
		' Well, well, well - looks like they made everything easy to setup through the
		' user interface except the description.  No matter - its easy if you know how
		' to work with the registry ...

		Dim ServicesKey As RegistryKey
		Dim ServiceNameKey As RegistryKey
		Try
			ServicesKey = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services", True)
			ServiceNameKey = ServicesKey.OpenSubKey(serviceName, True)
			If ServiceNameKey Is Nothing Then ServiceNameKey = ServicesKey.CreateSubKey(serviceName)
			ServiceNameKey.SetValue("Description", description)
		Finally
			If Not ServiceNameKey Is Nothing Then ServiceNameKey.Close()
			If Not ServicesKey Is Nothing Then ServicesKey.Close()
		End Try
	End Sub
End Class
