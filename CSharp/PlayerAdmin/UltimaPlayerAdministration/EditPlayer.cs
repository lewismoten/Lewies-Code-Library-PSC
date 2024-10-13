using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace UltimaPlayerAdministration
{
	/// <summary>
	/// Summary description for Player.
	/// </summary>
	public class EditPlayer : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button SavePlayer;
		private System.Windows.Forms.TextBox IPAddressText;
		private System.Windows.Forms.TextBox AccountText;
		private System.Windows.Forms.TextBox PlayerText;
		private System.Windows.Forms.Label IPAddressLabel;
		private System.Windows.Forms.Label AccountLabel;
		private System.Windows.Forms.Label PlayerLabel;
		public string Account;
		public string Player;
		public Guid PlayerID;
		public string IP;
		public string Action = "";
		public int Flag;
		private System.Windows.Forms.HScrollBar FlagScroll;
		private System.Windows.Forms.Label FlagLabel;
		private System.Windows.Forms.PictureBox FlagPicture;
		public System.Windows.Forms.ImageList FlagList;
		private System.Windows.Forms.LinkLabel DeletePlayerLink;
		private System.Windows.Forms.Button CancelEdit;
		private System.Windows.Forms.PictureBox pictureBox1;

		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public EditPlayer()
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
				if(components != null)
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
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(EditPlayer));
			this.SavePlayer = new System.Windows.Forms.Button();
			this.IPAddressText = new System.Windows.Forms.TextBox();
			this.AccountText = new System.Windows.Forms.TextBox();
			this.PlayerText = new System.Windows.Forms.TextBox();
			this.IPAddressLabel = new System.Windows.Forms.Label();
			this.AccountLabel = new System.Windows.Forms.Label();
			this.PlayerLabel = new System.Windows.Forms.Label();
			this.CancelEdit = new System.Windows.Forms.Button();
			this.FlagLabel = new System.Windows.Forms.Label();
			this.FlagPicture = new System.Windows.Forms.PictureBox();
			this.FlagScroll = new System.Windows.Forms.HScrollBar();
			this.DeletePlayerLink = new System.Windows.Forms.LinkLabel();
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.SuspendLayout();
			// 
			// SavePlayer
			// 
			this.SavePlayer.Location = new System.Drawing.Point(20, 160);
			this.SavePlayer.Name = "SavePlayer";
			this.SavePlayer.TabIndex = 13;
			this.SavePlayer.Text = "OK";
			this.SavePlayer.Click += new System.EventHandler(this.SavePlayer_Click);
			// 
			// IPAddressText
			// 
			this.IPAddressText.Location = new System.Drawing.Point(100, 100);
			this.IPAddressText.Name = "IPAddressText";
			this.IPAddressText.TabIndex = 12;
			this.IPAddressText.Text = "";
			// 
			// AccountText
			// 
			this.AccountText.Location = new System.Drawing.Point(100, 70);
			this.AccountText.Name = "AccountText";
			this.AccountText.TabIndex = 11;
			this.AccountText.Text = "";
			// 
			// PlayerText
			// 
			this.PlayerText.Location = new System.Drawing.Point(100, 40);
			this.PlayerText.Name = "PlayerText";
			this.PlayerText.TabIndex = 10;
			this.PlayerText.Text = "";
			// 
			// IPAddressLabel
			// 
			this.IPAddressLabel.AutoSize = true;
			this.IPAddressLabel.Location = new System.Drawing.Point(10, 100);
			this.IPAddressLabel.Name = "IPAddressLabel";
			this.IPAddressLabel.Size = new System.Drawing.Size(59, 16);
			this.IPAddressLabel.TabIndex = 9;
			this.IPAddressLabel.Text = "IP Address";
			// 
			// AccountLabel
			// 
			this.AccountLabel.AutoSize = true;
			this.AccountLabel.Location = new System.Drawing.Point(10, 70);
			this.AccountLabel.Name = "AccountLabel";
			this.AccountLabel.Size = new System.Drawing.Size(45, 16);
			this.AccountLabel.TabIndex = 8;
			this.AccountLabel.Text = "Account";
			// 
			// PlayerLabel
			// 
			this.PlayerLabel.AutoSize = true;
			this.PlayerLabel.Location = new System.Drawing.Point(10, 40);
			this.PlayerLabel.Name = "PlayerLabel";
			this.PlayerLabel.Size = new System.Drawing.Size(36, 16);
			this.PlayerLabel.TabIndex = 7;
			this.PlayerLabel.Text = "Player";
			// 
			// CancelEdit
			// 
			this.CancelEdit.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.CancelEdit.Location = new System.Drawing.Point(120, 160);
			this.CancelEdit.Name = "CancelEdit";
			this.CancelEdit.TabIndex = 14;
			this.CancelEdit.Text = "Cancel";
			this.CancelEdit.Click += new System.EventHandler(this.CancelEdit_Click);
			// 
			// FlagLabel
			// 
			this.FlagLabel.AutoSize = true;
			this.FlagLabel.Location = new System.Drawing.Point(10, 130);
			this.FlagLabel.Name = "FlagLabel";
			this.FlagLabel.Size = new System.Drawing.Size(26, 16);
			this.FlagLabel.TabIndex = 15;
			this.FlagLabel.Text = "Flag";
			// 
			// FlagPicture
			// 
			this.FlagPicture.BackColor = System.Drawing.Color.White;
			this.FlagPicture.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
			this.FlagPicture.Location = new System.Drawing.Point(100, 130);
			this.FlagPicture.Name = "FlagPicture";
			this.FlagPicture.Size = new System.Drawing.Size(18, 18);
			this.FlagPicture.TabIndex = 16;
			this.FlagPicture.TabStop = false;
			// 
			// FlagScroll
			// 
			this.FlagScroll.LargeChange = 1;
			this.FlagScroll.Location = new System.Drawing.Point(130, 130);
			this.FlagScroll.Name = "FlagScroll";
			this.FlagScroll.Size = new System.Drawing.Size(70, 17);
			this.FlagScroll.TabIndex = 17;
			this.FlagScroll.ValueChanged += new System.EventHandler(this.FlagScroll_ValueChanged);
			// 
			// DeletePlayerLink
			// 
			this.DeletePlayerLink.AutoSize = true;
			this.DeletePlayerLink.Location = new System.Drawing.Point(120, 10);
			this.DeletePlayerLink.Name = "DeletePlayerLink";
			this.DeletePlayerLink.Size = new System.Drawing.Size(72, 16);
			this.DeletePlayerLink.TabIndex = 19;
			this.DeletePlayerLink.TabStop = true;
			this.DeletePlayerLink.Text = "Delete Player";
			this.DeletePlayerLink.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.DeletePlayerLink_LinkClicked);
			// 
			// pictureBox1
			// 
			this.pictureBox1.Cursor = System.Windows.Forms.Cursors.Hand;
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(100, 10);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(14, 16);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
			this.pictureBox1.TabIndex = 21;
			this.pictureBox1.TabStop = false;
			this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
			// 
			// EditPlayer
			// 
			this.AcceptButton = this.SavePlayer;
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.CancelButton = this.CancelEdit;
			this.ClientSize = new System.Drawing.Size(212, 196);
			this.Controls.Add(this.pictureBox1);
			this.Controls.Add(this.DeletePlayerLink);
			this.Controls.Add(this.FlagScroll);
			this.Controls.Add(this.FlagPicture);
			this.Controls.Add(this.FlagLabel);
			this.Controls.Add(this.CancelEdit);
			this.Controls.Add(this.SavePlayer);
			this.Controls.Add(this.IPAddressText);
			this.Controls.Add(this.AccountText);
			this.Controls.Add(this.PlayerText);
			this.Controls.Add(this.IPAddressLabel);
			this.Controls.Add(this.AccountLabel);
			this.Controls.Add(this.PlayerLabel);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "EditPlayer";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
			this.Text = "Edit Player";
			this.TopMost = true;
			this.VisibleChanged += new System.EventHandler(this.EditPlayer_VisibleChanged);
			this.ResumeLayout(false);

		}
		#endregion

		private void SavePlayer_Click(object sender, System.EventArgs e)
		{
			Action = "Update";
			Player = PlayerText.Text;
			Account = AccountText.Text;
			IP = IPAddressText.Text;
			Close();
		}

		private void FlagScroll_ValueChanged(object sender, System.EventArgs e)
		{
			if(FlagScroll.Value == -1)
				FlagPicture.Image = null;
			else
				FlagPicture.Image = FlagList.Images[FlagScroll.Value];
			Flag = FlagScroll.Value;
		}

		private void EditPlayer_VisibleChanged(object sender, System.EventArgs e)
		{
			FlagScroll.Maximum = FlagList.Images.Count-1;
			FlagScroll.Minimum = -1;
			FlagScroll.Value = Flag;
			FlagScroll_ValueChanged(sender, e);

			PlayerText.Text = Player;
			AccountText.Text = Account;
			IPAddressText.Text = IP;
		}

		private void DeletePlayerLink_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			if(MessageBox.Show(this, 
				"You are about to delete this player.", 
				"Confirmation", 
				MessageBoxButtons.OKCancel,
				MessageBoxIcon.Warning,
				MessageBoxDefaultButton.Button2) == DialogResult.Cancel) return;
			Action = "Delete";
			Close();
		}

		private void CancelEdit_Click(object sender, System.EventArgs e)
		{
			Close();
		}

		private void pictureBox1_Click(object sender, System.EventArgs e)
		{
			DeletePlayerLink_LinkClicked(sender, null);
		}


	}
}
