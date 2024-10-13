Namespace Services
	Public MustInherit Class BaseService

		Protected _Service As Config.Alerts.Service
		Protected _DataProvider As DataStore.BaseDataStore

		Public Property Service() As Config.Alerts.Service
			Get
				Return _Service
			End Get
			Set(ByVal Value As Config.Alerts.Service)
				_Service = Value
			End Set
		End Property
		Public Property DataProvider() As DataStore.BaseDataStore
			Get
				Return _DataProvider
			End Get
			Set(ByVal Value As DataStore.BaseDataStore)
				_DataProvider = Value
			End Set
		End Property

		Public Sub CheckAlerts(ByVal pollingInterval As Double, ByVal Offset As Integer)
			For Each alert As Config.Alerts.Service.Alert In Me.Service.Alerts

				If Not (alert.FixedTime Is Nothing Or alert.FixedTime = "") Then

					' Parse Time of Day
					Dim Time As DateTime = CType(alert.FixedTime, DateTime)

					Dim Utc As DateTime = DateTime.Now.ToUniversalTime

					' Determine time today that message is sent
					Dim NextSending As New DateTime(Utc.Year, Utc.Month, Utc.Day, Time.Hour, Time.Minute, Time.Second)
					NextSending = NextSending.AddHours(Offset * -1)
					If Not NextSending.Day = Utc.Day Then
						NextSending.AddDays(Math.Sign(Offset))
					End If

					' if we are within the time range of sending the message ...
					If Math.Abs(Utc.Subtract(NextSending).TotalSeconds) < (pollingInterval / 1000) Then
						If SendAlert(alert) Then alert.LastSent = Utc
					End If

				Else
					' alert not sent within range of delivery cycle?
					If alert.LastSent.AddMinutes(alert.DeliveryCycle) < Now.ToUniversalTime Then
						' If we sent alert, update last sent time
						If SendAlert(alert) Then alert.LastSent = Now.ToUniversalTime
					End If
				End If
			Next
		End Sub

		Public MustOverride Function SendAlert(ByVal alert As Config.Alerts.Service.Alert) As Boolean

		Protected Overloads Function Compose(ByVal message As Config.Alerts.Service.Alert.MessageFormat, ByVal entriesSet As EntryData) As String
			Return Compose(message, entriesSet.Entries.Select())
		End Function
		Protected Overloads Function Compose(ByVal message As Config.Alerts.Service.Alert.MessageFormat, ByVal entriesTable As EntryData.EntriesDataTable) As String
			Return Compose(message, entriesTable.Select())
		End Function
		Protected Overloads Function Compose(ByVal message As Config.Alerts.Service.Alert.MessageFormat, ByVal entriesRows() As EntryData.EntriesRow) As String

			' create string builder
			Dim Text As New System.Text.StringBuilder("")

			' add header
			If Not message.Header Is Nothing Then
				Text.Append(message.Header.Content)
			End If

			' gather item templates
			Dim Item As String = message.Item.Content
			Dim AlternateItem As String = message.Item.Content
			If Not message.AlternateItem Is Nothing Then
				AlternateItem = message.AlternateItem.Content
			End If
			If AlternateItem = "" Then AlternateItem = Item

			Dim Alternate As Boolean = False

			Dim FormatProvider As New FormatProvider

			' loop through entries
			For Each Entry As EntryData.EntriesRow In entriesRows

				' display text
				If Alternate Then
					Text.Append(String.Format(FormatProvider, AlternateItem, Entry.ItemArray))
				Else
					Text.Append(String.Format(FormatProvider, Item, Entry.ItemArray))
				End If

				' toggle alternate item status
				Alternate = Not Alternate

			Next

			' tag on the footer
			If Not message.Footer Is Nothing Then
				Text.Append(message.Footer.Content)
			End If

			Return Text.ToString

		End Function

	End Class
End Namespace
