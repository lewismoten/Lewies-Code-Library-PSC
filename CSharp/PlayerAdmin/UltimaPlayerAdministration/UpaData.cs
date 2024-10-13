﻿//------------------------------------------------------------------------------
// <autogenerated>
//     This code was generated by a tool.
//     Runtime Version: 1.1.4322.573
//
//     Changes to this file may cause incorrect behavior and will be lost if 
//     the code is regenerated.
// </autogenerated>
//------------------------------------------------------------------------------

namespace UltimaPlayerAdministration {
    using System;
    using System.Data;
    using System.Xml;
    using System.Runtime.Serialization;
    
    
    [Serializable()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Diagnostics.DebuggerStepThrough()]
    [System.ComponentModel.ToolboxItem(true)]
    public class UpaData : DataSet {
        
        private PlayersDataTable tablePlayers;
        
        public UpaData() {
            this.InitClass();
            System.ComponentModel.CollectionChangeEventHandler schemaChangedHandler = new System.ComponentModel.CollectionChangeEventHandler(this.SchemaChanged);
            this.Tables.CollectionChanged += schemaChangedHandler;
            this.Relations.CollectionChanged += schemaChangedHandler;
        }
        
        protected UpaData(SerializationInfo info, StreamingContext context) {
            string strSchema = ((string)(info.GetValue("XmlSchema", typeof(string))));
            if ((strSchema != null)) {
                DataSet ds = new DataSet();
                ds.ReadXmlSchema(new XmlTextReader(new System.IO.StringReader(strSchema)));
                if ((ds.Tables["Players"] != null)) {
                    this.Tables.Add(new PlayersDataTable(ds.Tables["Players"]));
                }
                this.DataSetName = ds.DataSetName;
                this.Prefix = ds.Prefix;
                this.Namespace = ds.Namespace;
                this.Locale = ds.Locale;
                this.CaseSensitive = ds.CaseSensitive;
                this.EnforceConstraints = ds.EnforceConstraints;
                this.Merge(ds, false, System.Data.MissingSchemaAction.Add);
                this.InitVars();
            }
            else {
                this.InitClass();
            }
            this.GetSerializationData(info, context);
            System.ComponentModel.CollectionChangeEventHandler schemaChangedHandler = new System.ComponentModel.CollectionChangeEventHandler(this.SchemaChanged);
            this.Tables.CollectionChanged += schemaChangedHandler;
            this.Relations.CollectionChanged += schemaChangedHandler;
        }
        
        [System.ComponentModel.Browsable(false)]
        [System.ComponentModel.DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Content)]
        public PlayersDataTable Players {
            get {
                return this.tablePlayers;
            }
        }
        
        public override DataSet Clone() {
            UpaData cln = ((UpaData)(base.Clone()));
            cln.InitVars();
            return cln;
        }
        
        protected override bool ShouldSerializeTables() {
            return false;
        }
        
        protected override bool ShouldSerializeRelations() {
            return false;
        }
        
        protected override void ReadXmlSerializable(XmlReader reader) {
            this.Reset();
            DataSet ds = new DataSet();
            ds.ReadXml(reader);
            if ((ds.Tables["Players"] != null)) {
                this.Tables.Add(new PlayersDataTable(ds.Tables["Players"]));
            }
            this.DataSetName = ds.DataSetName;
            this.Prefix = ds.Prefix;
            this.Namespace = ds.Namespace;
            this.Locale = ds.Locale;
            this.CaseSensitive = ds.CaseSensitive;
            this.EnforceConstraints = ds.EnforceConstraints;
            this.Merge(ds, false, System.Data.MissingSchemaAction.Add);
            this.InitVars();
        }
        
        protected override System.Xml.Schema.XmlSchema GetSchemaSerializable() {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            this.WriteXmlSchema(new XmlTextWriter(stream, null));
            stream.Position = 0;
            return System.Xml.Schema.XmlSchema.Read(new XmlTextReader(stream), null);
        }
        
        internal void InitVars() {
            this.tablePlayers = ((PlayersDataTable)(this.Tables["Players"]));
            if ((this.tablePlayers != null)) {
                this.tablePlayers.InitVars();
            }
        }
        
        private void InitClass() {
            this.DataSetName = "UpaData";
            this.Prefix = "";
            this.Namespace = "";
            this.Locale = new System.Globalization.CultureInfo("en-US");
            this.CaseSensitive = false;
            this.EnforceConstraints = true;
            this.tablePlayers = new PlayersDataTable();
            this.Tables.Add(this.tablePlayers);
        }
        
        private bool ShouldSerializePlayers() {
            return false;
        }
        
        private void SchemaChanged(object sender, System.ComponentModel.CollectionChangeEventArgs e) {
            if ((e.Action == System.ComponentModel.CollectionChangeAction.Remove)) {
                this.InitVars();
            }
        }
        
        public delegate void PlayersRowChangeEventHandler(object sender, PlayersRowChangeEvent e);
        
        [System.Diagnostics.DebuggerStepThrough()]
        public class PlayersDataTable : DataTable, System.Collections.IEnumerable {
            
            private DataColumn columnPlayerID;
            
            private DataColumn columnPlayer;
            
            private DataColumn columnAccount;
            
            private DataColumn columnIP;
            
            private DataColumn columnCreated;
            
            private DataColumn columnModified;
            
            private DataColumn columnFlag;
            
            internal PlayersDataTable() : 
                    base("Players") {
                this.InitClass();
            }
            
            internal PlayersDataTable(DataTable table) : 
                    base(table.TableName) {
                if ((table.CaseSensitive != table.DataSet.CaseSensitive)) {
                    this.CaseSensitive = table.CaseSensitive;
                }
                if ((table.Locale.ToString() != table.DataSet.Locale.ToString())) {
                    this.Locale = table.Locale;
                }
                if ((table.Namespace != table.DataSet.Namespace)) {
                    this.Namespace = table.Namespace;
                }
                this.Prefix = table.Prefix;
                this.MinimumCapacity = table.MinimumCapacity;
                this.DisplayExpression = table.DisplayExpression;
            }
            
            [System.ComponentModel.Browsable(false)]
            public int Count {
                get {
                    return this.Rows.Count;
                }
            }
            
            internal DataColumn PlayerIDColumn {
                get {
                    return this.columnPlayerID;
                }
            }
            
            internal DataColumn PlayerColumn {
                get {
                    return this.columnPlayer;
                }
            }
            
            internal DataColumn AccountColumn {
                get {
                    return this.columnAccount;
                }
            }
            
            internal DataColumn IPColumn {
                get {
                    return this.columnIP;
                }
            }
            
            internal DataColumn CreatedColumn {
                get {
                    return this.columnCreated;
                }
            }
            
            internal DataColumn ModifiedColumn {
                get {
                    return this.columnModified;
                }
            }
            
            internal DataColumn FlagColumn {
                get {
                    return this.columnFlag;
                }
            }
            
            public PlayersRow this[int index] {
                get {
                    return ((PlayersRow)(this.Rows[index]));
                }
            }
            
            public event PlayersRowChangeEventHandler PlayersRowChanged;
            
            public event PlayersRowChangeEventHandler PlayersRowChanging;
            
            public event PlayersRowChangeEventHandler PlayersRowDeleted;
            
            public event PlayersRowChangeEventHandler PlayersRowDeleting;
            
            public void AddPlayersRow(PlayersRow row) {
                this.Rows.Add(row);
            }
            
            public PlayersRow AddPlayersRow(System.Guid PlayerID, string Player, string Account, string IP, System.DateTime Created, System.DateTime Modified, int Flag) {
                PlayersRow rowPlayersRow = ((PlayersRow)(this.NewRow()));
                rowPlayersRow.ItemArray = new object[] {
                        PlayerID,
                        Player,
                        Account,
                        IP,
                        Created,
                        Modified,
                        Flag};
                this.Rows.Add(rowPlayersRow);
                return rowPlayersRow;
            }
            
            public PlayersRow FindByPlayerID(System.Guid PlayerID) {
                return ((PlayersRow)(this.Rows.Find(new object[] {
                            PlayerID})));
            }
            
            public System.Collections.IEnumerator GetEnumerator() {
                return this.Rows.GetEnumerator();
            }
            
            public override DataTable Clone() {
                PlayersDataTable cln = ((PlayersDataTable)(base.Clone()));
                cln.InitVars();
                return cln;
            }
            
            protected override DataTable CreateInstance() {
                return new PlayersDataTable();
            }
            
            internal void InitVars() {
                this.columnPlayerID = this.Columns["PlayerID"];
                this.columnPlayer = this.Columns["Player"];
                this.columnAccount = this.Columns["Account"];
                this.columnIP = this.Columns["IP"];
                this.columnCreated = this.Columns["Created"];
                this.columnModified = this.Columns["Modified"];
                this.columnFlag = this.Columns["Flag"];
            }
            
            private void InitClass() {
                this.columnPlayerID = new DataColumn("PlayerID", typeof(System.Guid), null, System.Data.MappingType.Attribute);
                this.Columns.Add(this.columnPlayerID);
                this.columnPlayer = new DataColumn("Player", typeof(string), null, System.Data.MappingType.Attribute);
                this.Columns.Add(this.columnPlayer);
                this.columnAccount = new DataColumn("Account", typeof(string), null, System.Data.MappingType.Attribute);
                this.Columns.Add(this.columnAccount);
                this.columnIP = new DataColumn("IP", typeof(string), null, System.Data.MappingType.Attribute);
                this.Columns.Add(this.columnIP);
                this.columnCreated = new DataColumn("Created", typeof(System.DateTime), null, System.Data.MappingType.Attribute);
                this.Columns.Add(this.columnCreated);
                this.columnModified = new DataColumn("Modified", typeof(System.DateTime), null, System.Data.MappingType.Attribute);
                this.Columns.Add(this.columnModified);
                this.columnFlag = new DataColumn("Flag", typeof(int), null, System.Data.MappingType.Attribute);
                this.Columns.Add(this.columnFlag);
                this.Constraints.Add(new UniqueConstraint("PlayersPK", new DataColumn[] {
                                this.columnPlayerID}, true));
                this.columnPlayerID.AllowDBNull = false;
                this.columnPlayerID.Unique = true;
                this.columnPlayerID.Namespace = "";
                this.columnPlayer.Namespace = "";
                this.columnAccount.Namespace = "";
                this.columnIP.Namespace = "";
                this.columnCreated.Namespace = "";
                this.columnModified.Namespace = "";
                this.columnFlag.Namespace = "";
            }
            
            public PlayersRow NewPlayersRow() {
                return ((PlayersRow)(this.NewRow()));
            }
            
            protected override DataRow NewRowFromBuilder(DataRowBuilder builder) {
                return new PlayersRow(builder);
            }
            
            protected override System.Type GetRowType() {
                return typeof(PlayersRow);
            }
            
            protected override void OnRowChanged(DataRowChangeEventArgs e) {
                base.OnRowChanged(e);
                if ((this.PlayersRowChanged != null)) {
                    this.PlayersRowChanged(this, new PlayersRowChangeEvent(((PlayersRow)(e.Row)), e.Action));
                }
            }
            
            protected override void OnRowChanging(DataRowChangeEventArgs e) {
                base.OnRowChanging(e);
                if ((this.PlayersRowChanging != null)) {
                    this.PlayersRowChanging(this, new PlayersRowChangeEvent(((PlayersRow)(e.Row)), e.Action));
                }
            }
            
            protected override void OnRowDeleted(DataRowChangeEventArgs e) {
                base.OnRowDeleted(e);
                if ((this.PlayersRowDeleted != null)) {
                    this.PlayersRowDeleted(this, new PlayersRowChangeEvent(((PlayersRow)(e.Row)), e.Action));
                }
            }
            
            protected override void OnRowDeleting(DataRowChangeEventArgs e) {
                base.OnRowDeleting(e);
                if ((this.PlayersRowDeleting != null)) {
                    this.PlayersRowDeleting(this, new PlayersRowChangeEvent(((PlayersRow)(e.Row)), e.Action));
                }
            }
            
            public void RemovePlayersRow(PlayersRow row) {
                this.Rows.Remove(row);
            }
        }
        
        [System.Diagnostics.DebuggerStepThrough()]
        public class PlayersRow : DataRow {
            
            private PlayersDataTable tablePlayers;
            
            internal PlayersRow(DataRowBuilder rb) : 
                    base(rb) {
                this.tablePlayers = ((PlayersDataTable)(this.Table));
            }
            
            public System.Guid PlayerID {
                get {
                    return ((System.Guid)(this[this.tablePlayers.PlayerIDColumn]));
                }
                set {
                    this[this.tablePlayers.PlayerIDColumn] = value;
                }
            }
            
            public string Player {
                get {
                    try {
                        return ((string)(this[this.tablePlayers.PlayerColumn]));
                    }
                    catch (InvalidCastException e) {
                        throw new StrongTypingException("Cannot get value because it is DBNull.", e);
                    }
                }
                set {
                    this[this.tablePlayers.PlayerColumn] = value;
                }
            }
            
            public string Account {
                get {
                    try {
                        return ((string)(this[this.tablePlayers.AccountColumn]));
                    }
                    catch (InvalidCastException e) {
                        throw new StrongTypingException("Cannot get value because it is DBNull.", e);
                    }
                }
                set {
                    this[this.tablePlayers.AccountColumn] = value;
                }
            }
            
            public string IP {
                get {
                    try {
                        return ((string)(this[this.tablePlayers.IPColumn]));
                    }
                    catch (InvalidCastException e) {
                        throw new StrongTypingException("Cannot get value because it is DBNull.", e);
                    }
                }
                set {
                    this[this.tablePlayers.IPColumn] = value;
                }
            }
            
            public System.DateTime Created {
                get {
                    try {
                        return ((System.DateTime)(this[this.tablePlayers.CreatedColumn]));
                    }
                    catch (InvalidCastException e) {
                        throw new StrongTypingException("Cannot get value because it is DBNull.", e);
                    }
                }
                set {
                    this[this.tablePlayers.CreatedColumn] = value;
                }
            }
            
            public System.DateTime Modified {
                get {
                    try {
                        return ((System.DateTime)(this[this.tablePlayers.ModifiedColumn]));
                    }
                    catch (InvalidCastException e) {
                        throw new StrongTypingException("Cannot get value because it is DBNull.", e);
                    }
                }
                set {
                    this[this.tablePlayers.ModifiedColumn] = value;
                }
            }
            
            public int Flag {
                get {
                    try {
                        return ((int)(this[this.tablePlayers.FlagColumn]));
                    }
                    catch (InvalidCastException e) {
                        throw new StrongTypingException("Cannot get value because it is DBNull.", e);
                    }
                }
                set {
                    this[this.tablePlayers.FlagColumn] = value;
                }
            }
            
            public bool IsPlayerNull() {
                return this.IsNull(this.tablePlayers.PlayerColumn);
            }
            
            public void SetPlayerNull() {
                this[this.tablePlayers.PlayerColumn] = System.Convert.DBNull;
            }
            
            public bool IsAccountNull() {
                return this.IsNull(this.tablePlayers.AccountColumn);
            }
            
            public void SetAccountNull() {
                this[this.tablePlayers.AccountColumn] = System.Convert.DBNull;
            }
            
            public bool IsIPNull() {
                return this.IsNull(this.tablePlayers.IPColumn);
            }
            
            public void SetIPNull() {
                this[this.tablePlayers.IPColumn] = System.Convert.DBNull;
            }
            
            public bool IsCreatedNull() {
                return this.IsNull(this.tablePlayers.CreatedColumn);
            }
            
            public void SetCreatedNull() {
                this[this.tablePlayers.CreatedColumn] = System.Convert.DBNull;
            }
            
            public bool IsModifiedNull() {
                return this.IsNull(this.tablePlayers.ModifiedColumn);
            }
            
            public void SetModifiedNull() {
                this[this.tablePlayers.ModifiedColumn] = System.Convert.DBNull;
            }
            
            public bool IsFlagNull() {
                return this.IsNull(this.tablePlayers.FlagColumn);
            }
            
            public void SetFlagNull() {
                this[this.tablePlayers.FlagColumn] = System.Convert.DBNull;
            }
        }
        
        [System.Diagnostics.DebuggerStepThrough()]
        public class PlayersRowChangeEvent : EventArgs {
            
            private PlayersRow eventRow;
            
            private DataRowAction eventAction;
            
            public PlayersRowChangeEvent(PlayersRow row, DataRowAction action) {
                this.eventRow = row;
                this.eventAction = action;
            }
            
            public PlayersRow Row {
                get {
                    return this.eventRow;
                }
            }
            
            public DataRowAction Action {
                get {
                    return this.eventAction;
                }
            }
        }
    }
}
