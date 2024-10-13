Imports System.Xml.Serialization
Imports System.Data.SqlClient
Imports VMS.Framework
Imports System.IO

Public Class Xml
	Inherits DataStore.BaseDataStore

	Public Overrides Sub CacheEntries(ByVal EntryData As Framework.EntryData)

		Dim Cache As New VMS.Framework.EntryData
		Try
			' load existing xml
			If File.Exists(DataStore("Path")) Then
				Cache.ReadXml(DataStore("Path"))
			End If

			' add non-dupes
			For Each EntryRow As EntryData.EntriesRow In EntryData.Entries

				With EntryRow
					If Cache.Entries.FindBy_DatePathVersionAction(._Date, .Path, .Version, .Action) Is Nothing Then
						Cache.Entries.Rows.Add(.ItemArray)
					End If
				End With
			Next

			' save xml
			Cache.WriteXml(DataStore("Path"))

		Finally
			Cache.Dispose()
		End Try

	End Sub

	Public Overrides Function FilterEntries(ByVal alert As Framework.Config.Alerts.Service.Alert) As Framework.EntryData

		Dim Cache As New EntryData
		Dim Results As New EntryData
		Try

			If Not File.Exists(DataStore("Path")) Then Return Results

			' load existing xml
			Cache.ReadXml(DataStore("Path"))

			' prepare filter
			Dim filter As String = alert.Filter.Content
			filter = filter.Replace("@LastAlerted", "'" & alert.LastSent & "'")

			' apply filter
			Dim Rows() As DataRow = Cache.Entries.Select(filter)

			For Each Row As DataRow In Rows
				Results.Entries.ImportRow(Row)
			Next

			' return rows
			Return Results
		Finally
			Results.Dispose()
			Cache.Dispose()
		End Try

	End Function
End Class
