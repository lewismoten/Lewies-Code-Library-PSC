Imports System.ComponentModel

<RunInstaller(True)> Public Class ProjectInstaller
	Inherits System.Configuration.Install.Installer

	Private Sub ProjectInstaller_AfterInstall(ByVal sender As Object, ByVal e As System.Configuration.Install.InstallEventArgs) Handles MyBase.AfterInstall
		Dim Form As New Install
		Form.ShowDialog()
		Form = Nothing
	End Sub
End Class
