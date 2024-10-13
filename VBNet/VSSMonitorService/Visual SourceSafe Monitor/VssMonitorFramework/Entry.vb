Imports System.Text.RegularExpressions
Imports System.Text.RegularExpressions.RegexOptions
Imports System.Configuration.ConfigurationSettings

Public Class Entry

	Private _Match As Match
	Private _Users As Config.Users

	Private _path As String = ""
	Private _version As Integer = -1
	Private _user As String = ""
	Private _date As String = ""
	Private _action As String = ""
	Private _event As String = ""
	Private _comment As String = ""

	Public Shared Function ToDataset(ByVal journal As String) As EntryData

		Dim RawEntries As MatchCollection = Entry.ParseEntries(journal)
		Dim EntryData As New EntryData

		' Loop through each entry
		For Each RawEntry As Match In RawEntries
			With New Entry(RawEntry, Config.Users.Instance)
				EntryData.Entries.AddEntriesRow(Now.ToUniversalTime, .Date, .Path, .Version, .Action, .User, .Event, .Comment)
			End With
		Next

		Return EntryData

	End Function
	Public Shared Function ParseEntries(ByVal journal As String) As MatchCollection

		' Example of entry within journal:

		'$/NEWWAVE/Applications/Diagnostics/DataAccessLogic
		'Version: 2
		'User: Lewis Moten     Date:  1/08/04  Time: 11:20p
		'Assemblies.vb added
		'Comment: Imported project files into Visual SourceSafe.

		Dim pattern As String = ""

		' File/Project affected
		pattern &= "(\$/.*)\r\n"

		' Version
		pattern &= "Version:\s+(\d+)\r\n"

		' User
		pattern &= "User:\s+(.+?)\s+"

		' Date
		pattern &= "Date:\s+((\d+)/(\d+)/(\d+))\s+"

		' Time
		pattern &= "Time:\s+((\d+):(\d+)([ap]))\r\n"

		' Action
		pattern &= "([^\r]*)\r\n"

		' User Comments (optional)
		pattern &= "(Comment:\s+(([\s\S]+?)\r\n))?"

		' End of entry
		pattern &= "\r\n"

		Dim Expression As New Regex(pattern, Multiline)

		Return Expression.Matches(journal)

	End Function

	Public Sub New(ByVal rawEntry As Match, ByVal users As Config.Users)
		_Match = rawEntry
		_Users = users
	End Sub

	Public ReadOnly Property Path() As String
		Get
			If _path = "" Then
				_path = _Match.Groups(1).Value
			End If
			Return _path
		End Get
	End Property

	Public ReadOnly Property Version() As Integer
		Get
			If _version = -1 Then
				_version = CType(_Match.Groups(2).Value, Integer)
			End If
			Return _version
		End Get
	End Property

	Public ReadOnly Property User() As String
		Get
			If _user = "" Then
				_user = _Match.Groups(3).Value
			End If
			Return _user
		End Get
	End Property

	Private Function Y2KFix(ByVal year As Integer) As Integer
		If year < 85 Then
			Return year + 2000
		Else
			Return year + 1900
		End If
	End Function

	Public ReadOnly Property [Date]() As DateTime
		Get
			If Not _date = "" Then Return CType(_date, DateTime)

			' Dates may become problematic with timezones
			' when daylight savings time goes into effect
			' at different times in different regions of
			' the world.

			Dim D1 As Integer = CType(_Match.Groups(5).Value, Integer)
			Dim D2 As Integer = CType(_Match.Groups(6).Value, Integer)
			Dim D3 As Integer = CType(_Match.Groups(7).Value, Integer)
			Dim Hour As Integer = CType(_Match.Groups(9).Value, Integer)
			Dim Minute As Integer = CType(_Match.Groups(10).Value, Integer)

			' convert hour to "military time"
			If Hour = 12 Then Hour = 0
			If _Match.Groups(11).Value = "p" Then Hour += 12

			' get default date format
			Dim DateFormat As String = _Users.DefaultDateFormat
			Dim Offset As Integer = _Users.DefaultTimzeZone

			' Default date format to US date
			If DateFormat Is Nothing Or DateFormat = "" Then DateFormat = "MM/DD/YY"

			' find user
			For Each user As Config.Users.User In _Users.Users
				If user.Name = MyClass.User Then
					' get users date format
					If Not (user.DateFormat Is Nothing Or user.DateFormat = "") Then
						DateFormat = user.DateFormat
					End If
					If Math.Abs(user.TimeZone) < 24 Then Offset = user.TimeZone
				End If
			Next

			Dim ReturnDate As DateTime

			' return new date based on format provided
			Select Case DateFormat
				Case "MM/DD/YY"				' US Date
					ReturnDate = New Date(Y2KFix(D3), D1, D2, Hour, Minute, 0)
				Case "DD/MM/YY"				' Europe
					ReturnDate = New Date(Y2KFix(D3), D2, D1, Hour, Minute, 0)
				Case "YY/MM/DD"				' Programmer
					ReturnDate = New Date(Y2KFix(D1), D2, D3, Hour, Minute, 0)
				Case Else
					Throw New Exception("Invalid date format specified.  Expecting MM/DD/YY, DD/MM/YY, or YY/MM/DD.")
			End Select

			' apply timezone adjustment (converts to GMT)
			' NOTE: Daily Savings may cause problems.
			' other countries may start/stop time zones on different dates
			' some do not observe daylight savings at all.
			ReturnDate = ReturnDate.AddHours(Offset * -1)

			' cache date
			_date = ReturnDate.ToString

			Return ReturnDate
		End Get
	End Property

	Public ReadOnly Property Action() As String
		Get
			If _action = "" Then _action = _Match.Groups(12).Value
			Return _action
		End Get
	End Property

	Public ReadOnly Property [Event]() As String
		Get
			If Not _event = "" Then Return _event

			If Action = "Checked in" Then
				_event = "Checked in"
			ElseIf Action = "Rolled back" Then
				_event = "Rolled back"
			ElseIf Action.IndexOf(" copied from $/") > 0 Then
				_event = "Copied"
			ElseIf Action.IndexOf(" moved to $/") > 0 Then
				_event = "Moved"
			ElseIf Action.IndexOf(" renamed to ") > 0 Then
				_event = "Renamed"
			ElseIf Action.IndexOf(" shared from $/") > 0 Then
				_event = "Shared"
			ElseIf Action.StartsWith("Labeled ") Then
				' HACK: Creating a file called "Labled .txt" will return "Labeled"
				_event = "Labeled"
			ElseIf Action.EndsWith(" added") Then
				_event = "Added"
			ElseIf Action.EndsWith(" branched") Then
				_event = "Branched"
			ElseIf Action.EndsWith(" created") Then
				_event = "Created"
			ElseIf Action.EndsWith(" deleted") Then
				_event = "Deleted"
			ElseIf Action.EndsWith(" destroyed") Then
				_event = "Destroyed"
			ElseIf Action.EndsWith(" purged") Then
				_event = "Purged"
			ElseIf Action.EndsWith(" recovered") Then
				_event = "Recovered"
			Else
				_event = "Unknown"
			End If

			Return _event

		End Get
	End Property

	Public ReadOnly Property Comment() As String
		Get
			If _comment = "" Then _comment = _Match.Groups(14).Value
			Return _comment
		End Get
	End Property

End Class