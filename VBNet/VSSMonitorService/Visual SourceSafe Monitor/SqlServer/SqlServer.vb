Imports System.Xml.Serialization
Imports System.Data.SqlClient
Imports VMS.Framework

' Class used to store and retrieve entries from an SQL server database.
' You must install the Database.SQL file that came with this source code
' in order for this to work properly.

Public Class SqlServer
	Inherits VMS.Framework.DataStore.BaseDataStore

	Public Overrides Sub CacheEntries(ByVal EntryData As EntryData)

		Dim Connection As New SqlConnection
		Dim Command As New SqlCommand
		Dim Adapter As New SqlDataAdapter

		Try
			Connection.ConnectionString = ConnectionString

			' setup command to insert events
			Command.Connection = Connection
			Command.CommandType = CommandType.StoredProcedure
			Command.CommandText = "[dbo].[spr_CacheEntry]"

			' When adding parameters, map each one to a field in the datatable
			Command.Parameters.Add("@Logged", SqlDbType.SmallDateTime, 4, "Logged")
			Command.Parameters.Add("@Date", SqlDbType.SmallDateTime, 4, "Date")
			Command.Parameters.Add("@Path", SqlDbType.VarChar, 256, "Path")
			Command.Parameters.Add("@Version", SqlDbType.Int, 4, "Version")
			Command.Parameters.Add("@Action", SqlDbType.VarChar, 128, "Action")
			Command.Parameters.Add("@User", SqlDbType.VarChar, 32, "User")
			Command.Parameters.Add("@Event", SqlDbType.VarChar, 16, "Event")
			Command.Parameters.Add("@Comment", SqlDbType.Text, Int32.MaxValue, "Comment")

			' Add all events with one update
			Adapter.InsertCommand = Command
			Adapter.Update(EntryData.Entries)
		Finally
			If Not Adapter Is Nothing Then
				Adapter.Dispose()
				Adapter = Nothing
			End If

			If Not Command Is Nothing Then
				Command.Dispose()
				Command = Nothing
			End If

			If Not Connection Is Nothing Then
				If Connection.State = ConnectionState.Open Then
					Connection.Close()
				End If
				Connection.Dispose()
				Connection = Nothing
			End If
		End Try

	End Sub

	<XmlIgnore()> Private ReadOnly Property ConnectionString() As String
		Get
			' Create a connection string based on properties provided

			Dim Server As String = DataStore("Server").Replace("=", "==")
			Dim Catalog As String = DataStore("Catalog").Replace("=", "==")
			Dim Username As String = DataStore("Username").Replace("=", "==")
			Dim Password As String = DataStore("Password").Replace("=", "==")
			Dim Trusted As Boolean = False
			If Not DataStore("Trusted") = "" Then
				Trusted = CType(DataStore("Trusted"), Boolean)
			End If

			Dim Text As New System.Text.StringBuilder
			Text.Append("Persist Security Info=False;")
			If Trusted Then
				Text.Append("Integrated Security=SSPI;")
			Else
				Text.Append("Password=" & Password & ";")
				Text.Append("User ID=" & Username & ";")
			End If
			Text.Append("Initial Catalog=" & Catalog & ";")
			Text.Append("server=" & Server & ";")
			Text.Append("Application Name=VssMonitorService;")

			Return Text.ToString

		End Get
	End Property

	Public Overrides Function FilterEntries(ByVal alert As Config.Alerts.Service.Alert) As EntryData

		' Filter Entries expects an SQL Query that will return the 
		' following fields and data type in the order they are
		' specified

		' Logged	smalldatetime
		' Date		smalldatetime
		' Path		varchar(256)
		' Version	int
		' Action	varchar(128)
		' User		varchar(32)
		' Event		varchar(16)
		' Comment	text

		' In addition, you may use a varialbe called @LastAlerted
		' that will be replaced with the date that the last email
		' was sent for this particular alert.

		Dim Connection As New SqlConnection
		Dim Command As New SqlCommand
		Dim Adapter As New SqlDataAdapter
		Dim EntryData As New EntryData
		Try
			Connection.ConnectionString = ConnectionString

			Command.Connection = Connection
			Command.CommandType = CommandType.Text

			Command.CommandText = alert.Filter.Content
			Command.Parameters.Add("@LastAlerted", alert.LastSent)

			Adapter.SelectCommand = Command
			Adapter.Fill(EntryData.Entries)
			Return EntryData
		Finally
			If Not Adapter Is Nothing Then
				Adapter.Dispose()
				Adapter = Nothing
			End If

			If Not Command Is Nothing Then
				Command.Dispose()
				Command = Nothing
			End If

			If Not Connection Is Nothing Then
				If Connection.State = ConnectionState.Open Then
					Connection.Close()
				End If
				Connection.Dispose()
				Connection = Nothing
			End If
		End Try

	End Function
End Class