using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Web;

using System.Runtime.InteropServices;

namespace UltimaPlayerAdministration
{
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Panel SearchPanel;
		private System.Windows.Forms.Label SearchLabel;
		private System.Windows.Forms.TextBox SearchText;
		private System.Windows.Forms.Panel DataEntryPanel;
		private System.Windows.Forms.Label PlayerLabel;
		private System.Windows.Forms.Label AccountLabel;
		private System.Windows.Forms.Label IPAddressLabel;
		private System.Windows.Forms.TextBox PlayerText;
		private System.Windows.Forms.TextBox AccountText;
		private System.Windows.Forms.TextBox IPAddressText;
		private System.Windows.Forms.Button AddPlayer;
		private System.Windows.Forms.ListView PlayerList;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.ColumnHeader columnHeader2;
		private System.Windows.Forms.ColumnHeader columnHeader3;
		private System.ComponentModel.IContainer components;
		private System.Windows.Forms.Label MatchingResultLabel;
		private System.Windows.Forms.Label MatchingResultCount;
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem Exit;
		private System.Windows.Forms.MenuItem Export;
		private System.Windows.Forms.MenuItem Import;
		private System.Windows.Forms.OpenFileDialog openDialog;
		private System.Windows.Forms.SaveFileDialog saveDialog;
		private Guid SelectedPlayerID = Guid.Empty;
		private System.Windows.Forms.MenuItem menuItem4;
		private System.Windows.Forms.MenuItem menuItem3;
		private System.Windows.Forms.ImageList FlagList;
		private System.Windows.Forms.MenuItem SendDonation;
		private System.Windows.Forms.MenuItem AuthorsWebsite;

		private UpaData data = null;


		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(Form1));
			this.SearchPanel = new System.Windows.Forms.Panel();
			this.MatchingResultCount = new System.Windows.Forms.Label();
			this.MatchingResultLabel = new System.Windows.Forms.Label();
			this.SearchText = new System.Windows.Forms.TextBox();
			this.SearchLabel = new System.Windows.Forms.Label();
			this.DataEntryPanel = new System.Windows.Forms.Panel();
			this.AddPlayer = new System.Windows.Forms.Button();
			this.IPAddressText = new System.Windows.Forms.TextBox();
			this.AccountText = new System.Windows.Forms.TextBox();
			this.PlayerText = new System.Windows.Forms.TextBox();
			this.IPAddressLabel = new System.Windows.Forms.Label();
			this.AccountLabel = new System.Windows.Forms.Label();
			this.PlayerLabel = new System.Windows.Forms.Label();
			this.PlayerList = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.Export = new System.Windows.Forms.MenuItem();
			this.Import = new System.Windows.Forms.MenuItem();
			this.Exit = new System.Windows.Forms.MenuItem();
			this.openDialog = new System.Windows.Forms.OpenFileDialog();
			this.saveDialog = new System.Windows.Forms.SaveFileDialog();
			this.SendDonation = new System.Windows.Forms.MenuItem();
			this.menuItem4 = new System.Windows.Forms.MenuItem();
			this.menuItem3 = new System.Windows.Forms.MenuItem();
			this.AuthorsWebsite = new System.Windows.Forms.MenuItem();
			this.FlagList = new System.Windows.Forms.ImageList(this.components);
			this.SearchPanel.SuspendLayout();
			this.DataEntryPanel.SuspendLayout();
			this.SuspendLayout();
			// 
			// SearchPanel
			// 
			this.SearchPanel.Controls.Add(this.MatchingResultCount);
			this.SearchPanel.Controls.Add(this.MatchingResultLabel);
			this.SearchPanel.Controls.Add(this.SearchText);
			this.SearchPanel.Controls.Add(this.SearchLabel);
			this.SearchPanel.Dock = System.Windows.Forms.DockStyle.Top;
			this.SearchPanel.Location = new System.Drawing.Point(0, 0);
			this.SearchPanel.Name = "SearchPanel";
			this.SearchPanel.Size = new System.Drawing.Size(472, 40);
			this.SearchPanel.TabIndex = 0;
			// 
			// MatchingResultCount
			// 
			this.MatchingResultCount.AutoSize = true;
			this.MatchingResultCount.Location = new System.Drawing.Point(290, 10);
			this.MatchingResultCount.Name = "MatchingResultCount";
			this.MatchingResultCount.Size = new System.Drawing.Size(10, 16);
			this.MatchingResultCount.TabIndex = 4;
			this.MatchingResultCount.Text = "0";
			// 
			// MatchingResultLabel
			// 
			this.MatchingResultLabel.AutoSize = true;
			this.MatchingResultLabel.Location = new System.Drawing.Point(180, 10);
			this.MatchingResultLabel.Name = "MatchingResultLabel";
			this.MatchingResultLabel.Size = new System.Drawing.Size(94, 16);
			this.MatchingResultLabel.TabIndex = 3;
			this.MatchingResultLabel.Text = "Matching Results:";
			// 
			// SearchText
			// 
			this.SearchText.Location = new System.Drawing.Point(60, 10);
			this.SearchText.Name = "SearchText";
			this.SearchText.TabIndex = 1;
			this.SearchText.Text = "";
			this.SearchText.KeyUp += new System.Windows.Forms.KeyEventHandler(this.SearchText_KeyUp);
			// 
			// SearchLabel
			// 
			this.SearchLabel.AutoSize = true;
			this.SearchLabel.BackColor = System.Drawing.SystemColors.Control;
			this.SearchLabel.Location = new System.Drawing.Point(10, 10);
			this.SearchLabel.Name = "SearchLabel";
			this.SearchLabel.Size = new System.Drawing.Size(43, 16);
			this.SearchLabel.TabIndex = 0;
			this.SearchLabel.Text = "Search:";
			// 
			// DataEntryPanel
			// 
			this.DataEntryPanel.Controls.Add(this.AddPlayer);
			this.DataEntryPanel.Controls.Add(this.IPAddressText);
			this.DataEntryPanel.Controls.Add(this.AccountText);
			this.DataEntryPanel.Controls.Add(this.PlayerText);
			this.DataEntryPanel.Controls.Add(this.IPAddressLabel);
			this.DataEntryPanel.Controls.Add(this.AccountLabel);
			this.DataEntryPanel.Controls.Add(this.PlayerLabel);
			this.DataEntryPanel.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.DataEntryPanel.Location = new System.Drawing.Point(0, 326);
			this.DataEntryPanel.Name = "DataEntryPanel";
			this.DataEntryPanel.Size = new System.Drawing.Size(472, 60);
			this.DataEntryPanel.TabIndex = 1;
			// 
			// AddPlayer
			// 
			this.AddPlayer.Location = new System.Drawing.Point(380, 30);
			this.AddPlayer.Name = "AddPlayer";
			this.AddPlayer.TabIndex = 6;
			this.AddPlayer.Text = "Add Player";
			this.AddPlayer.Click += new System.EventHandler(this.AddPlayer_Click);
			// 
			// IPAddressText
			// 
			this.IPAddressText.Location = new System.Drawing.Point(250, 30);
			this.IPAddressText.Name = "IPAddressText";
			this.IPAddressText.TabIndex = 5;
			this.IPAddressText.Text = "";
			// 
			// AccountText
			// 
			this.AccountText.Location = new System.Drawing.Point(130, 30);
			this.AccountText.Name = "AccountText";
			this.AccountText.TabIndex = 4;
			this.AccountText.Text = "";
			// 
			// PlayerText
			// 
			this.PlayerText.Location = new System.Drawing.Point(10, 30);
			this.PlayerText.Name = "PlayerText";
			this.PlayerText.TabIndex = 3;
			this.PlayerText.Text = "";
			// 
			// IPAddressLabel
			// 
			this.IPAddressLabel.AutoSize = true;
			this.IPAddressLabel.Location = new System.Drawing.Point(250, 10);
			this.IPAddressLabel.Name = "IPAddressLabel";
			this.IPAddressLabel.Size = new System.Drawing.Size(59, 16);
			this.IPAddressLabel.TabIndex = 2;
			this.IPAddressLabel.Text = "IP Address";
			// 
			// AccountLabel
			// 
			this.AccountLabel.AutoSize = true;
			this.AccountLabel.Location = new System.Drawing.Point(130, 10);
			this.AccountLabel.Name = "AccountLabel";
			this.AccountLabel.Size = new System.Drawing.Size(45, 16);
			this.AccountLabel.TabIndex = 1;
			this.AccountLabel.Text = "Account";
			// 
			// PlayerLabel
			// 
			this.PlayerLabel.AutoSize = true;
			this.PlayerLabel.Location = new System.Drawing.Point(10, 10);
			this.PlayerLabel.Name = "PlayerLabel";
			this.PlayerLabel.Size = new System.Drawing.Size(36, 16);
			this.PlayerLabel.TabIndex = 0;
			this.PlayerLabel.Text = "Player";
			// 
			// PlayerList
			// 
			this.PlayerList.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						 this.columnHeader1,
																						 this.columnHeader2,
																						 this.columnHeader3});
			this.PlayerList.Dock = System.Windows.Forms.DockStyle.Fill;
			this.PlayerList.FullRowSelect = true;
			this.PlayerList.HideSelection = false;
			this.PlayerList.Location = new System.Drawing.Point(0, 40);
			this.PlayerList.Name = "PlayerList";
			this.PlayerList.Size = new System.Drawing.Size(472, 286);
			this.PlayerList.SmallImageList = this.FlagList;
			this.PlayerList.TabIndex = 2;
			this.PlayerList.View = System.Windows.Forms.View.Details;
			this.PlayerList.DoubleClick += new System.EventHandler(this.PlayerList_DoubleClick);
			this.PlayerList.SelectedIndexChanged += new System.EventHandler(this.PlayerList_SelectedIndexChanged);
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "Player";
			this.columnHeader1.Width = 120;
			// 
			// columnHeader2
			// 
			this.columnHeader2.Text = "Account";
			this.columnHeader2.Width = 120;
			// 
			// columnHeader3
			// 
			this.columnHeader3.Text = "IP Address";
			this.columnHeader3.Width = 120;
			// 
			// mainMenu1
			// 
			this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.menuItem1,
																					  this.menuItem3});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.Export,
																					  this.Import,
																					  this.menuItem4,
																					  this.Exit});
			this.menuItem1.Text = "File";
			// 
			// Export
			// 
			this.Export.Index = 0;
			this.Export.Text = "Export ...";
			this.Export.Click += new System.EventHandler(this.Export_Click);
			// 
			// Import
			// 
			this.Import.Index = 1;
			this.Import.Text = "Import ...";
			this.Import.Click += new System.EventHandler(this.Import_Click);
			// 
			// Exit
			// 
			this.Exit.Index = 3;
			this.Exit.Text = "Exit";
			this.Exit.Click += new System.EventHandler(this.Exit_Click);
			// 
			// SendDonation
			// 
			this.SendDonation.Index = 0;
			this.SendDonation.Text = "Send Donation";
			this.SendDonation.Click += new System.EventHandler(this.SendDonation_Click);
			// 
			// menuItem4
			// 
			this.menuItem4.Index = 2;
			this.menuItem4.Text = "-";
			// 
			// menuItem3
			// 
			this.menuItem3.Index = 1;
			this.menuItem3.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.SendDonation,
																					  this.AuthorsWebsite});
			this.menuItem3.Text = "Misc.";
			// 
			// AuthorsWebsite
			// 
			this.AuthorsWebsite.Index = 1;
			this.AuthorsWebsite.Text = "Authors Website";
			this.AuthorsWebsite.Click += new System.EventHandler(this.AuthorsWebsite_Click);
			// 
			// FlagList
			// 
			this.FlagList.ImageSize = new System.Drawing.Size(16, 16);
			this.FlagList.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("FlagList.ImageStream")));
			this.FlagList.TransparentColor = System.Drawing.Color.Transparent;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(472, 386);
			this.Controls.Add(this.PlayerList);
			this.Controls.Add(this.DataEntryPanel);
			this.Controls.Add(this.SearchPanel);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Menu = this.mainMenu1;
			this.Name = "Form1";
			this.Text = "Ultima Player Administration";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.SearchPanel.ResumeLayout(false);
			this.DataEntryPanel.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private string FileName
		{
			get
			{
				return System.IO.Path.GetDirectoryName(Application.UserAppDataPath) + "\\UpaData.xml";
			}
		}
		private void Form1_Load(object sender, System.EventArgs e)
		{
			data = new UpaData();

			if(System.IO.File.Exists(FileName))
				data.ReadXml(FileName);

			PopulateList();
		}

		private void AddPlayer_Click(object sender, System.EventArgs e)
		{
			if(PlayerText.Text.Trim() == "" && AccountText.Text.Trim() == "" && IPAddressText.Text.Trim() == "")
			{
				MessageBox.Show(this, "Please enter some player information before adding them to the list.", "Missing Data");
				return;
			}
			data.Players.AddPlayersRow(Guid.NewGuid(), PlayerText.Text, AccountText.Text, IPAddressText.Text, DateTime.Now.ToUniversalTime(), DateTime.Now.ToUniversalTime(), -1);
			data.WriteXml(FileName);

			PlayerText.Text = "";
			AccountText.Text = "";
			IPAddressText.Text = "";

			PopulateList();
		}

		private void PopulateList(int topIndex)
		{
			PopulateList();
			if(topIndex < PlayerList.Items.Count)
				for(int i = 0; i < PlayerList.Items.Count; i++)
				{
					PlayerList.Items[i].EnsureVisible();
					if(topIndex == PlayerList.TopItem.Index) break;
				}
			if(PlayerList.SelectedItems.Count != 0)
				PlayerList.SelectedItems[0].EnsureVisible();
		}
		private void PopulateList()
		{
			PlayerList.Items.Clear();
			data.Players.DefaultView.Sort = "Player ASC";

			MatchingResultCount.Text = data.Players.DefaultView.Count.ToString();

			IEnumerator e = data.Players.DefaultView.GetEnumerator();

			while(e.MoveNext())
			{
				UpaData.PlayersRow Player = (UpaData.PlayersRow)((DataRowView)e.Current).Row;
				PlayerList.Items.Add(new ListViewItem(new string[]{Player.Player, Player.Account, Player.IP}));
					
				if(Player.PlayerID.Equals(SelectedPlayerID))
					PlayerList.Items[PlayerList.Items.Count-1].Selected = true;
				PlayerList.Items[PlayerList.Items.Count-1].ImageIndex = Player.Flag;
			}

			if(PlayerList.SelectedItems.Count != 0)
				PlayerList.SelectedItems[0].EnsureVisible();

		}

		private void SearchText_KeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
		{
			string query = SearchText.Text;
			string filter = "Player LIKE '%{0}%' OR Account LIKE '%{0}%' OR IP LIKE '%{0}%'";

			if(query == "")
				data.Players.DefaultView.RowFilter = "";
			else
				data.Players.DefaultView.RowFilter = string.Format(filter, query);

			PopulateList();
	
		}

		private void PlayerList_DoubleClick(object sender, System.EventArgs e)
		{
			if(PlayerList.SelectedItems.Count == 0) return;

			ListViewItem Player = PlayerList.SelectedItems[0];

			UpaData.PlayersRow row = (UpaData.PlayersRow)data.Players.DefaultView[Player.Index].Row;

			EditPlayer form = new EditPlayer();
			form.Player = row.Player;
			form.Account = row.Account;
			form.IP = row.IP;
			form.Flag = row.Flag;
			form.FlagList = FlagList;

			form.ShowDialog(this);

			if(form.Action == "Delete")
			{
				row.Delete();
				data.AcceptChanges();
				data.WriteXml(FileName);
				PopulateList(PlayerList.TopItem.Index);
			}
			else if(form.Action == "Update")
			{
				row.Account = form.Account;
				row.Player = form.Player;
				row.IP = form.IP;
				row.Flag = form.Flag;
				row.Modified = DateTime.Now.ToUniversalTime();
				data.AcceptChanges();
				data.WriteXml(FileName);
				PopulateList(PlayerList.TopItem.Index);
			}
			form.Dispose();

		}

		private void Exit_Click(object sender, System.EventArgs e)
		{
			Application.Exit();
		}

		private void Export_Click(object sender, System.EventArgs e)
		{
			saveDialog.AddExtension = true;
			saveDialog.DefaultExt = "xml";
			saveDialog.Filter = "Xml Data Files (*.xml)|*.xml";
			saveDialog.OverwritePrompt = true;
			saveDialog.Title = "Select location to save exported data";
			if(saveDialog.ShowDialog(this) == DialogResult.Cancel)
				return;
			data.WriteXml(saveDialog.FileName);
		}

		private void Import_Click(object sender, System.EventArgs e)
		{
			openDialog.AddExtension = true;
			openDialog.DefaultExt = "xml";
			openDialog.Filter = "Xml Data Files (*.xml)|*.xml";
			openDialog.Title = "Select Xml data file to import player information";
			if(openDialog.ShowDialog(this) == DialogResult.Cancel)
				return;
			
			UpaData import = new UpaData();
			import.ReadXml(openDialog.FileName);

			int added = 0;
			int updated = 0;
			int ignored = 0;

			for(int i = 0; i < import.Players.Count; i++)
			{
				UpaData.PlayersRow newPlayer = import.Players[i];
				UpaData.PlayersRow oldPlayer = data.Players.FindByPlayerID(newPlayer.PlayerID);

				if(oldPlayer == null)
				{
					// add new rows
					oldPlayer = data.Players.NewPlayersRow();
					oldPlayer.ItemArray = newPlayer.ItemArray;
					data.Players.AddPlayersRow(oldPlayer);
					added++;
				}
				else if(oldPlayer.Modified < newPlayer.Modified)
				{
					// update exiting rows
					oldPlayer.IP = newPlayer.IP;
					oldPlayer.Player = newPlayer.Player;
					oldPlayer.Account = newPlayer.Account;
					oldPlayer.Modified = newPlayer.Modified;
					updated++;
				}
				else
				{
					ignored++;
				}
			}
			import.Dispose();
			data.AcceptChanges();
			data.WriteXml(FileName);

			PopulateList();

			string message = "New Records: {0}\r\nModified Records: {1}\r\nOld Records: {2}\r\nTotal: {3}";
			MessageBox.Show(this, string.Format(message, added, updated, ignored, added + updated + ignored), "Import Complete");
		}

		private void PlayerList_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if(PlayerList.SelectedItems.Count != 0)
				SelectedPlayerID = ((UpaData.PlayersRow)data.Players.DefaultView[PlayerList.SelectedItems[0].Index].Row).PlayerID;
		}

		private void AuthorsWebsite_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("http://www.lewismoten.com");
		}

		private void SendDonation_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("https://www.paypal.com/xclick/business=paypal%40magic-in-a-package.com&item_name=" + HttpUtility.UrlEncode(Application.ProductName) + "&item_number=" + HttpUtility.UrlEncode(Application.ProductVersion) + "&no_note=1&tax=0&currency_code=USD");
		}

	}
}
