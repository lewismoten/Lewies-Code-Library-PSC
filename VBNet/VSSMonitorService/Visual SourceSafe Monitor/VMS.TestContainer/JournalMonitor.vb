Imports System.ServiceProcess
Imports System.IO
Imports System.Timers
Imports System.Reflection
Imports System.Xml.Serialization
Imports System.Configuration.ConfigurationSettings
Imports VMS.Framework
Public Class JournalMonitor
	Inherits System.ServiceProcess.ServiceBase
	Protected Settings As Config.Settings
	Protected Alerts As Config.Alerts
	Protected Users As Config.Users
	Protected JournalPath As String
	Protected PollingInterval As Double
	Protected DataProvider As DataStore.BaseDataStore
	Protected ServiceProviders() As VMS.Framework.Services.BaseService

#Region " Component Designer generated code "

	Public Sub New()
		MyBase.New()

		' This call is required by the Component Designer.
		InitializeComponent()

		' Add any initialization after the InitializeComponent() call

	End Sub

	'UserService overrides dispose to clean up the component list.
	Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
		If disposing Then
			If Not (components Is Nothing) Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(disposing)
	End Sub

	' The main entry point for the process
	<MTAThread()> _
	Shared Sub Main()
		Dim ServicesToRun() As System.ServiceProcess.ServiceBase

		' More than one NT Service may run within the same process. To add
		' another service to this process, change the following line to
		' create a second service object. For example,
		'
		'   ServicesToRun = New System.ServiceProcess.ServiceBase () {New Service1, New MySecondUserService}
		'
		ServicesToRun = New System.ServiceProcess.ServiceBase() {New JournalMonitor}

		System.ServiceProcess.ServiceBase.Run(ServicesToRun)
	End Sub

	'Required by the Component Designer
	Private components As System.ComponentModel.IContainer

	' NOTE: The following procedure is required by the Component Designer
	' It can be modified using the Component Designer.  
	' Do not modify it using the code editor.
	Friend WithEvents Timer1 As System.Timers.Timer
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Me.Timer1 = New System.Timers.Timer
		CType(Me.Timer1, System.ComponentModel.ISupportInitialize).BeginInit()
		'
		'Timer1
		'
		Me.Timer1.AutoReset = False
		Me.Timer1.Interval = 60000
		'
		'JournalMonitor
		'
		Me.CanPauseAndContinue = True
		Me.CanShutdown = True
		Me.ServiceName = "Visual SourceSafe Journal Monitor"
		CType(Me.Timer1, System.ComponentModel.ISupportInitialize).EndInit()

	End Sub

#End Region

	Public Enum eventTypes
		OnStart
		OnStop
		OnContinue
		OnPause
		OnShutDown
	End Enum
	Public Sub TestEvent(ByVal name As eventTypes)
		Select Case name
			Case eventTypes.OnStart : OnStart(Nothing)
			Case eventTypes.OnStop : OnStop()
			Case eventTypes.OnContinue : OnContinue()
			Case eventTypes.OnPause : OnPause()
			Case eventTypes.OnShutDown : OnShutdown()
		End Select
	End Sub

	Protected Overrides Sub OnStart(ByVal args() As String)
		LoadSettings()
		StartService()
	End Sub

	Protected Overrides Sub OnStop()
		StopService()
		Settings.Save()
	End Sub

	Protected Overrides Sub OnContinue()
		LoadSettings()
		StartService()
	End Sub

	Protected Overrides Sub OnPause()
		StopService()
		Settings.Save()
	End Sub

	Protected Overrides Sub OnShutdown()
		StopService()
		Settings.Save()
	End Sub

	Private Sub LoadSettings()

		' This method reads all the configuration information
		' and caches them as objects.  it also loads up each
		' specified "type" for services and the datastore
		' to make certain everything exists before officially
		' starting the service.

		Settings = Settings.Instance()
		Users = Users.Instance
		Alerts = Alerts.Instance
		JournalPath = AppSettings("JournalPath")
		PollingInterval = CType(AppSettings("Interval"), Double) * 1000 * 60

		Dim DataStore As Config.DataStore = DataStore.Instance
		Dim DataProviderType As Type = System.Type.GetType(DataStore.Type, True)
		DataProvider = Activator.CreateInstance(DataProviderType)
		DataProvider.DataStore = DataStore

		Dim ServiceProviders(Alerts.Services.Length - 1) As VMS.Framework.Services.BaseService
		For i As Integer = 0 To Alerts.Services.Length - 1
			Dim Service As Config.Alerts.Service = Alerts.Services(i)
			Dim ServiceType As Type = System.Type.GetType(Service.Type, True)
			Dim ServiceProvider As VMS.Framework.Services.BaseService
			ServiceProvider = Activator.CreateInstance(ServiceType)
			ServiceProvider.Service = Service
			ServiceProvider.DataProvider = DataProvider
			ServiceProviders(i) = ServiceProvider
		Next
		Me.ServiceProviders = ServiceProviders

	End Sub

	Private Sub RunCheck()
		Try
			CheckJournal()
			CheckAlerts()
		Catch ex As Exception
			EventLog.WriteEntry(ServiceName, ex.GetBaseException.Message & vbCrLf & vbCrLf & ex.StackTrace, EventLogEntryType.Error)
		End Try
	End Sub

	Private Sub CheckJournal()

		Dim LastModified As DateTime = Settings.LastModified
		Dim FileModified As DateTime = DateTime.MinValue
		Dim FileLength As Long = 0
		Dim LastLength As Long = Settings.Length


		' See if file exists
		Dim FileExists As Boolean = File.Exists(JournalPath)
		If FileExists Then
			FileModified = File.GetLastWriteTime(JournalPath)
			FileLength = FileLen(JournalPath)
		End If

		' Log event if existence does not match previouse state.
		If Not FileExists = Settings.FileExists Then

			Settings.FileExists = FileExists

			' Some configurations may watch file on another
			' server.  If that server goes down for a task
			' such as reboot, we wait until it is detected
			' again and log those events.

			If FileExists Then
				EventLog.WriteEntry(ServiceName, "The file has been found." & vbCrLf & JournalPath, EventLogEntryType.SuccessAudit)
			Else
				EventLog.WriteEntry(ServiceName, "The file is not available.  Make sure that the file exists and that this service has permission to access it." & vbCrLf & JournalPath, EventLogEntryType.FailureAudit)
			End If

		End If

		If Not FileExists Then Return

		' File wasn't modified since last check?  Don't bother ...
		If FileModified <= LastModified Then Return

		' Filesize didn't change?  Don't bother ...
		If FileLength = LastLength Then Return

		Dim JournalText As String

		If FileLength < LastLength Then
			' Filesize shrunk ... let's read all of it.
			JournalText = Common.ReadTextFile(JournalPath)
		Else
			' Filesize grew ... let's read the new stuff
			JournalText = Common.ReadTextFile(JournalPath, LastLength)
		End If

		' Parse entries to a dataset
		Dim JournalData As EntryData = Entry.ToDataset(JournalText)
		JournalText = Nothing

		' Save entries to data provider
		DataProvider.CacheEntries(JournalData)

		Settings.LastModified = FileModified
		Settings.Length = FileLength

	End Sub

	Private Sub CheckAlerts()
		' Loop through all services
		For Each ServiceProvider As VMS.Framework.Services.BaseService In ServiceProviders
			' Tell services to check there alerts
			' (actually, just a base service object checking this info)
			ServiceProvider.CheckAlerts(PollingInterval, Alerts.Offset)
		Next
	End Sub
	Private Sub StartService()
		Timer1.Interval = PollingInterval
		Timer1.AutoReset = True
		Timer1.Start()
	End Sub

	Private Sub StopService()
		Timer1.Stop()
	End Sub

	Private Sub Timer1_Elapsed(ByVal sender As Object, ByVal e As System.Timers.ElapsedEventArgs) Handles Timer1.Elapsed
		RunCheck()
	End Sub

End Class
