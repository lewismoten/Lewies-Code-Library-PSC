using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace Mp3EditID3v1
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label SongTitleLabel;
		private System.Windows.Forms.TextBox SongTitleText;
		private System.Windows.Forms.Label ArtistLabel;
		private System.Windows.Forms.TextBox ArtistText;
		private System.Windows.Forms.Label AlbumLabel;
		private System.Windows.Forms.TextBox AlbumText;
		private System.Windows.Forms.Label YearLabel;
		private System.Windows.Forms.TextBox YearText;
		private System.Windows.Forms.Label TrackLabel;
		private System.Windows.Forms.TextBox TrackText;
		private System.Windows.Forms.Label GenreLabel;
		private System.Windows.Forms.ComboBox GenreList;
		private System.Windows.Forms.Label CommentLabel;
		private System.Windows.Forms.TextBox CommentText;
		private System.Windows.Forms.MenuItem menuItem9;
		private System.Windows.Forms.MenuItem FileMenuItem;
		private System.Windows.Forms.MenuItem ID3WebsiteMenuItem;
		private System.Windows.Forms.MenuItem AuthorsWebsiteMenuItem;
		private System.Windows.Forms.MenuItem LoadMenuItem;
		private System.Windows.Forms.MenuItem SaveMenuItem;
		private System.Windows.Forms.MenuItem ExitMenuItem;
		private System.Windows.Forms.MenuItem ID3v1_0MenuItem;
		private System.Windows.Forms.MenuItem ID3v1_1MenuItem;
		private System.Windows.Forms.MenuItem HelpMenuItem;
		private System.Windows.Forms.MainMenu MainMenu;
		private System.Windows.Forms.OpenFileDialog OpenFileDialog;
		private System.Windows.Forms.SaveFileDialog SaveFileDialog;
		private System.Windows.Forms.Label NewTagLabel;
		private string _fileName = "";
		private System.Windows.Forms.MenuItem VersionMenuItem;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

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
			this.SongTitleLabel = new System.Windows.Forms.Label();
			this.SongTitleText = new System.Windows.Forms.TextBox();
			this.ArtistLabel = new System.Windows.Forms.Label();
			this.ArtistText = new System.Windows.Forms.TextBox();
			this.AlbumLabel = new System.Windows.Forms.Label();
			this.AlbumText = new System.Windows.Forms.TextBox();
			this.YearLabel = new System.Windows.Forms.Label();
			this.YearText = new System.Windows.Forms.TextBox();
			this.TrackLabel = new System.Windows.Forms.Label();
			this.TrackText = new System.Windows.Forms.TextBox();
			this.GenreLabel = new System.Windows.Forms.Label();
			this.GenreList = new System.Windows.Forms.ComboBox();
			this.CommentLabel = new System.Windows.Forms.Label();
			this.CommentText = new System.Windows.Forms.TextBox();
			this.MainMenu = new System.Windows.Forms.MainMenu();
			this.FileMenuItem = new System.Windows.Forms.MenuItem();
			this.LoadMenuItem = new System.Windows.Forms.MenuItem();
			this.SaveMenuItem = new System.Windows.Forms.MenuItem();
			this.menuItem9 = new System.Windows.Forms.MenuItem();
			this.ExitMenuItem = new System.Windows.Forms.MenuItem();
			this.VersionMenuItem = new System.Windows.Forms.MenuItem();
			this.ID3v1_0MenuItem = new System.Windows.Forms.MenuItem();
			this.ID3v1_1MenuItem = new System.Windows.Forms.MenuItem();
			this.HelpMenuItem = new System.Windows.Forms.MenuItem();
			this.AuthorsWebsiteMenuItem = new System.Windows.Forms.MenuItem();
			this.ID3WebsiteMenuItem = new System.Windows.Forms.MenuItem();
			this.OpenFileDialog = new System.Windows.Forms.OpenFileDialog();
			this.SaveFileDialog = new System.Windows.Forms.SaveFileDialog();
			this.NewTagLabel = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// SongTitleLabel
			// 
			this.SongTitleLabel.AutoSize = true;
			this.SongTitleLabel.Location = new System.Drawing.Point(10, 10);
			this.SongTitleLabel.Name = "SongTitleLabel";
			this.SongTitleLabel.Size = new System.Drawing.Size(55, 16);
			this.SongTitleLabel.TabIndex = 0;
			this.SongTitleLabel.Text = "Song Title";
			// 
			// SongTitleText
			// 
			this.SongTitleText.Location = new System.Drawing.Point(10, 30);
			this.SongTitleText.MaxLength = 30;
			this.SongTitleText.Name = "SongTitleText";
			this.SongTitleText.Size = new System.Drawing.Size(340, 20);
			this.SongTitleText.TabIndex = 1;
			this.SongTitleText.Text = "";
			// 
			// ArtistLabel
			// 
			this.ArtistLabel.AutoSize = true;
			this.ArtistLabel.Location = new System.Drawing.Point(10, 60);
			this.ArtistLabel.Name = "ArtistLabel";
			this.ArtistLabel.Size = new System.Drawing.Size(30, 16);
			this.ArtistLabel.TabIndex = 5;
			this.ArtistLabel.Text = "Artist";
			// 
			// ArtistText
			// 
			this.ArtistText.Location = new System.Drawing.Point(10, 80);
			this.ArtistText.MaxLength = 30;
			this.ArtistText.Name = "ArtistText";
			this.ArtistText.Size = new System.Drawing.Size(340, 20);
			this.ArtistText.TabIndex = 6;
			this.ArtistText.Text = "";
			// 
			// AlbumLabel
			// 
			this.AlbumLabel.AutoSize = true;
			this.AlbumLabel.Location = new System.Drawing.Point(10, 110);
			this.AlbumLabel.Name = "AlbumLabel";
			this.AlbumLabel.Size = new System.Drawing.Size(36, 16);
			this.AlbumLabel.TabIndex = 7;
			this.AlbumLabel.Text = "Album";
			// 
			// AlbumText
			// 
			this.AlbumText.Location = new System.Drawing.Point(10, 130);
			this.AlbumText.MaxLength = 30;
			this.AlbumText.Name = "AlbumText";
			this.AlbumText.Size = new System.Drawing.Size(340, 20);
			this.AlbumText.TabIndex = 8;
			this.AlbumText.Text = "";
			// 
			// YearLabel
			// 
			this.YearLabel.AutoSize = true;
			this.YearLabel.Location = new System.Drawing.Point(10, 210);
			this.YearLabel.Name = "YearLabel";
			this.YearLabel.Size = new System.Drawing.Size(28, 16);
			this.YearLabel.TabIndex = 9;
			this.YearLabel.Text = "Year";
			// 
			// YearText
			// 
			this.YearText.Location = new System.Drawing.Point(10, 230);
			this.YearText.MaxLength = 4;
			this.YearText.Name = "YearText";
			this.YearText.Size = new System.Drawing.Size(30, 20);
			this.YearText.TabIndex = 10;
			this.YearText.Text = "";
			this.YearText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			// 
			// TrackLabel
			// 
			this.TrackLabel.AutoSize = true;
			this.TrackLabel.Location = new System.Drawing.Point(190, 210);
			this.TrackLabel.Name = "TrackLabel";
			this.TrackLabel.Size = new System.Drawing.Size(32, 16);
			this.TrackLabel.TabIndex = 11;
			this.TrackLabel.Text = "Track";
			// 
			// TrackText
			// 
			this.TrackText.Location = new System.Drawing.Point(190, 230);
			this.TrackText.MaxLength = 3;
			this.TrackText.Name = "TrackText";
			this.TrackText.Size = new System.Drawing.Size(30, 20);
			this.TrackText.TabIndex = 12;
			this.TrackText.Text = "";
			this.TrackText.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.TrackText.Validating += new System.ComponentModel.CancelEventHandler(this.TrackText_Validating);
			// 
			// GenreLabel
			// 
			this.GenreLabel.AutoSize = true;
			this.GenreLabel.Location = new System.Drawing.Point(60, 210);
			this.GenreLabel.Name = "GenreLabel";
			this.GenreLabel.Size = new System.Drawing.Size(36, 16);
			this.GenreLabel.TabIndex = 13;
			this.GenreLabel.Text = "Genre";
			// 
			// GenreList
			// 
			this.GenreList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.GenreList.Location = new System.Drawing.Point(60, 230);
			this.GenreList.Name = "GenreList";
			this.GenreList.Size = new System.Drawing.Size(120, 21);
			this.GenreList.TabIndex = 14;
			// 
			// CommentLabel
			// 
			this.CommentLabel.AutoSize = true;
			this.CommentLabel.Location = new System.Drawing.Point(10, 160);
			this.CommentLabel.Name = "CommentLabel";
			this.CommentLabel.Size = new System.Drawing.Size(53, 16);
			this.CommentLabel.TabIndex = 15;
			this.CommentLabel.Text = "Comment";
			// 
			// CommentText
			// 
			this.CommentText.Location = new System.Drawing.Point(10, 180);
			this.CommentText.MaxLength = 30;
			this.CommentText.Name = "CommentText";
			this.CommentText.Size = new System.Drawing.Size(340, 20);
			this.CommentText.TabIndex = 16;
			this.CommentText.Text = "";
			// 
			// MainMenu
			// 
			this.MainMenu.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					 this.FileMenuItem,
																					 this.VersionMenuItem,
																					 this.HelpMenuItem});
			// 
			// FileMenuItem
			// 
			this.FileMenuItem.Index = 0;
			this.FileMenuItem.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.LoadMenuItem,
																						 this.SaveMenuItem,
																						 this.menuItem9,
																						 this.ExitMenuItem});
			this.FileMenuItem.Text = "File";
			this.FileMenuItem.Select += new System.EventHandler(this.FileMenuItem_Select);
			// 
			// LoadMenuItem
			// 
			this.LoadMenuItem.Index = 0;
			this.LoadMenuItem.Text = "Load ...";
			this.LoadMenuItem.Click += new System.EventHandler(this.LoadMenuItem_Click);
			// 
			// SaveMenuItem
			// 
			this.SaveMenuItem.Index = 1;
			this.SaveMenuItem.Text = "Save";
			this.SaveMenuItem.Click += new System.EventHandler(this.SaveMenuItem_Click);
			// 
			// menuItem9
			// 
			this.menuItem9.Index = 2;
			this.menuItem9.Text = "-";
			// 
			// ExitMenuItem
			// 
			this.ExitMenuItem.Index = 3;
			this.ExitMenuItem.Text = "Exit";
			this.ExitMenuItem.Click += new System.EventHandler(this.ExitMenuItem_Click);
			// 
			// VersionMenuItem
			// 
			this.VersionMenuItem.Index = 1;
			this.VersionMenuItem.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																							this.ID3v1_0MenuItem,
																							this.ID3v1_1MenuItem});
			this.VersionMenuItem.Text = "Version";
			// 
			// ID3v1_0MenuItem
			// 
			this.ID3v1_0MenuItem.Index = 0;
			this.ID3v1_0MenuItem.Text = "ID3v1.0";
			this.ID3v1_0MenuItem.Click += new System.EventHandler(this.ID3v1_0MenuItem_Click);
			// 
			// ID3v1_1MenuItem
			// 
			this.ID3v1_1MenuItem.Checked = true;
			this.ID3v1_1MenuItem.Index = 1;
			this.ID3v1_1MenuItem.Text = "ID3v1.1";
			this.ID3v1_1MenuItem.Click += new System.EventHandler(this.ID3v1_1MenuItem_Click);
			// 
			// HelpMenuItem
			// 
			this.HelpMenuItem.Index = 2;
			this.HelpMenuItem.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																						 this.AuthorsWebsiteMenuItem,
																						 this.ID3WebsiteMenuItem});
			this.HelpMenuItem.Text = "Help";
			// 
			// AuthorsWebsiteMenuItem
			// 
			this.AuthorsWebsiteMenuItem.Index = 0;
			this.AuthorsWebsiteMenuItem.Text = "Authors Website";
			this.AuthorsWebsiteMenuItem.Click += new System.EventHandler(this.AuthorsWebsiteMenuItem_Click);
			// 
			// ID3WebsiteMenuItem
			// 
			this.ID3WebsiteMenuItem.Index = 1;
			this.ID3WebsiteMenuItem.Text = "ID3 Website";
			this.ID3WebsiteMenuItem.Click += new System.EventHandler(this.ID3WebsiteMenuItem_Click);
			// 
			// OpenFileDialog
			// 
			this.OpenFileDialog.DefaultExt = "mp3";
			this.OpenFileDialog.Filter = "MP3 Files (*.MP3)|*.MP3|All files (*.*)|*.*";
			this.OpenFileDialog.Title = "Open MP3 File";
			// 
			// NewTagLabel
			// 
			this.NewTagLabel.AutoSize = true;
			this.NewTagLabel.Location = new System.Drawing.Point(300, 230);
			this.NewTagLabel.Name = "NewTagLabel";
			this.NewTagLabel.Size = new System.Drawing.Size(51, 16);
			this.NewTagLabel.TabIndex = 17;
			this.NewTagLabel.Text = "(new tag)";
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(362, 266);
			this.Controls.Add(this.NewTagLabel);
			this.Controls.Add(this.CommentText);
			this.Controls.Add(this.CommentLabel);
			this.Controls.Add(this.GenreList);
			this.Controls.Add(this.GenreLabel);
			this.Controls.Add(this.TrackText);
			this.Controls.Add(this.TrackLabel);
			this.Controls.Add(this.YearText);
			this.Controls.Add(this.YearLabel);
			this.Controls.Add(this.AlbumText);
			this.Controls.Add(this.AlbumLabel);
			this.Controls.Add(this.ArtistText);
			this.Controls.Add(this.ArtistLabel);
			this.Controls.Add(this.SongTitleText);
			this.Controls.Add(this.SongTitleLabel);
			this.Menu = this.MainMenu;
			this.Name = "Form1";
			this.Text = "ID3v1 Tag Editor";
			this.Load += new System.EventHandler(this.Form1_Load);
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

		private void ID3v1_0MenuItem_Click(object sender, System.EventArgs e)
		{
			ID3v1_0MenuItem.Checked = true;
			ID3v1_1MenuItem.Checked = false;

			this.TrackLabel.Visible = false;
			this.TrackText.Visible = false;
			this.CommentText.MaxLength = 30;
		}

		private void ID3v1_1MenuItem_Click(object sender, System.EventArgs e)
		{
			ID3v1_0MenuItem.Checked = false;
			ID3v1_1MenuItem.Checked = true;

			this.TrackLabel.Visible = true;
			this.TrackText.Visible = true;
			this.CommentText.MaxLength = 28;
		}

		private void ExitMenuItem_Click(object sender, System.EventArgs e)
		{
			Application.Exit();
		}

		private void TrackText_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			if(int.Parse(TrackText.Text) != 0 && int.Parse(TrackText.Text) < 255) return;
			e.Cancel = true;
			MessageBox.Show("Track must be a number between 1 and 254.");
		}

		protected string FileName
		{
			get
			{
				return _fileName;
			}
			set
			{
				_fileName = value;
				Text = System.IO.Path.GetFileName(_fileName) + " - ID3v1 Tag Editor";
			}
		}
		private void LoadMenuItem_Click(object sender, System.EventArgs e)
		{
			if(DialogResult.Cancel == OpenFileDialog.ShowDialog(this)) return;
			
			FileName = OpenFileDialog.FileName;

			ID3v1 ID3 = new ID3v1();
			NewTagLabel.Visible = !ID3.TagExists(OpenFileDialog.FileName);
			ID3.Read(OpenFileDialog.FileName);
			SongTitleText.Text = ID3.SongTitle;
			ArtistText.Text = ID3.Artist;
			AlbumText.Text = ID3.Album;
			YearText.Text = ID3.Year;
			GenreList.SelectedIndex = ID3.Genre;

			if(ID3.Track == 0)
				ID3v1_0MenuItem_Click(this, e);
			else
				ID3v1_1MenuItem_Click(this, e);
			
			CommentText.Text = ID3.Comment;
			TrackText.Text = ID3.Track.ToString();
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			GenreList.DataSource = ID3v1.Genres();

			Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("Mp3EditID3v1.App.ico"); 
			System.Drawing.Icon icon = new System.Drawing.Icon(stream); 
			this.Icon = icon;

			Splash splash = new Splash();
			splash.Icon = icon;
			splash.ShowDialog(this);
			splash.Dispose();
			splash = null;
		}

		private void SaveMenuItem_Click(object sender, System.EventArgs e)
		{
			ID3v1 ID3 = new ID3v1();
			ID3.SongTitle = SongTitleText.Text;
			ID3.Artist = ArtistText.Text;
			ID3.Album = AlbumText.Text;
			ID3.Year = YearText.Text;
			ID3.Genre = (byte)GenreList.SelectedIndex;
			ID3.Comment = CommentText.Text;
			if(ID3v1_1MenuItem.Checked)
				ID3.Track = Convert.ToByte(TrackText.Text, 10);
			ID3.Write(FileName);
		}

		private void FileMenuItem_Select(object sender, System.EventArgs e)
		{
			SaveMenuItem.Enabled = FileName != "";
		}

		private void AuthorsWebsiteMenuItem_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("http://www.lewismoten.com");
		}

		private void ID3WebsiteMenuItem_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("http://www.id3.org");
		}
	}
}
