Namespace DataStore
	Public MustInherit Class BaseDataStore
		Inherits Config.BaseConfig
		Protected _DataStore As Config.DataStore

		' Method for caching entries into datastore
		Public MustOverride Sub CacheEntries(ByVal EntryData As EntryData)

		' method for filtering entries within datastore for reporting
		Public MustOverride Function FilterEntries(ByVal alert As Config.Alerts.Service.Alert) As EntryData

		' datastore property
		Public Property DataStore() As Config.DataStore
			Get
				Return _DataStore
			End Get
			Set(ByVal Value As Config.DataStore)
				_DataStore = Value
			End Set
		End Property
	End Class
End Namespace
