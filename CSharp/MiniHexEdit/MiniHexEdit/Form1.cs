using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Text;

namespace MiniHexEdit
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.VScrollBar vScrollBar1;
		private System.Windows.Forms.ListView listView1;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.ColumnHeader columnHeader2;
		private System.Windows.Forms.ColumnHeader columnHeader3;
		private System.Windows.Forms.MainMenu mainMenu1;
		private System.Windows.Forms.MenuItem menuItem1;
		private System.Windows.Forms.MenuItem OpenFileMenuItem;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
		private System.Windows.Forms.GroupBox BytesGroup;
		private System.Windows.Forms.Panel BytesPanel;
		private MiniHexEdit.ByteEdit byteEdit16;
		private MiniHexEdit.ByteEdit byteEdit15;
		private MiniHexEdit.ByteEdit byteEdit14;
		private MiniHexEdit.ByteEdit byteEdit13;
		private MiniHexEdit.ByteEdit byteEdit12;
		private MiniHexEdit.ByteEdit byteEdit11;
		private MiniHexEdit.ByteEdit byteEdit10;
		private MiniHexEdit.ByteEdit byteEdit9;
		private MiniHexEdit.ByteEdit byteEdit8;
		private MiniHexEdit.ByteEdit byteEdit7;
		private MiniHexEdit.ByteEdit byteEdit6;
		private MiniHexEdit.ByteEdit byteEdit5;
		private MiniHexEdit.ByteEdit byteEdit4;
		private MiniHexEdit.ByteEdit byteEdit3;
		private MiniHexEdit.ByteEdit byteEdit2;
		private MiniHexEdit.ByteEdit byteEdit1;
		private System.Windows.Forms.Button UpdateBytes;
		private System.Windows.Forms.MenuItem ExitProgram;
		private System.Windows.Forms.MenuItem AuthorsSiteMenuItem;
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
			System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem(new string[] {
																													 "00000000h",
																													 "4C-65-77-69-73-4D-6F-74-65-6E-2E-63-6F-6D-00-00",
																													 "LewisMoten.com.."}, -1);
			this.BytesGroup = new System.Windows.Forms.GroupBox();
			this.UpdateBytes = new System.Windows.Forms.Button();
			this.BytesPanel = new System.Windows.Forms.Panel();
			this.byteEdit16 = new MiniHexEdit.ByteEdit();
			this.byteEdit15 = new MiniHexEdit.ByteEdit();
			this.byteEdit14 = new MiniHexEdit.ByteEdit();
			this.byteEdit13 = new MiniHexEdit.ByteEdit();
			this.byteEdit12 = new MiniHexEdit.ByteEdit();
			this.byteEdit11 = new MiniHexEdit.ByteEdit();
			this.byteEdit10 = new MiniHexEdit.ByteEdit();
			this.byteEdit9 = new MiniHexEdit.ByteEdit();
			this.byteEdit8 = new MiniHexEdit.ByteEdit();
			this.byteEdit7 = new MiniHexEdit.ByteEdit();
			this.byteEdit6 = new MiniHexEdit.ByteEdit();
			this.byteEdit5 = new MiniHexEdit.ByteEdit();
			this.byteEdit4 = new MiniHexEdit.ByteEdit();
			this.byteEdit3 = new MiniHexEdit.ByteEdit();
			this.byteEdit2 = new MiniHexEdit.ByteEdit();
			this.byteEdit1 = new MiniHexEdit.ByteEdit();
			this.label5 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.vScrollBar1 = new System.Windows.Forms.VScrollBar();
			this.listView1 = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader2 = new System.Windows.Forms.ColumnHeader();
			this.columnHeader3 = new System.Windows.Forms.ColumnHeader();
			this.mainMenu1 = new System.Windows.Forms.MainMenu();
			this.menuItem1 = new System.Windows.Forms.MenuItem();
			this.OpenFileMenuItem = new System.Windows.Forms.MenuItem();
			this.AuthorsSiteMenuItem = new System.Windows.Forms.MenuItem();
			this.ExitProgram = new System.Windows.Forms.MenuItem();
			this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
			this.BytesGroup.SuspendLayout();
			this.BytesPanel.SuspendLayout();
			this.SuspendLayout();
			// 
			// BytesGroup
			// 
			this.BytesGroup.Controls.Add(this.UpdateBytes);
			this.BytesGroup.Controls.Add(this.BytesPanel);
			this.BytesGroup.Controls.Add(this.label5);
			this.BytesGroup.Controls.Add(this.label4);
			this.BytesGroup.Controls.Add(this.label3);
			this.BytesGroup.Controls.Add(this.label2);
			this.BytesGroup.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.BytesGroup.Location = new System.Drawing.Point(0, 115);
			this.BytesGroup.Name = "BytesGroup";
			this.BytesGroup.Size = new System.Drawing.Size(582, 280);
			this.BytesGroup.TabIndex = 33;
			this.BytesGroup.TabStop = false;
			this.BytesGroup.Text = "Bytes";
			// 
			// UpdateBytes
			// 
			this.UpdateBytes.Location = new System.Drawing.Point(10, 240);
			this.UpdateBytes.Name = "UpdateBytes";
			this.UpdateBytes.Size = new System.Drawing.Size(50, 23);
			this.UpdateBytes.TabIndex = 90;
			this.UpdateBytes.Text = "Update";
			this.UpdateBytes.Visible = false;
			this.UpdateBytes.Click += new System.EventHandler(this.UpdateBytes_Click);
			// 
			// BytesPanel
			// 
			this.BytesPanel.Controls.Add(this.byteEdit16);
			this.BytesPanel.Controls.Add(this.byteEdit15);
			this.BytesPanel.Controls.Add(this.byteEdit14);
			this.BytesPanel.Controls.Add(this.byteEdit13);
			this.BytesPanel.Controls.Add(this.byteEdit12);
			this.BytesPanel.Controls.Add(this.byteEdit11);
			this.BytesPanel.Controls.Add(this.byteEdit10);
			this.BytesPanel.Controls.Add(this.byteEdit9);
			this.BytesPanel.Controls.Add(this.byteEdit8);
			this.BytesPanel.Controls.Add(this.byteEdit7);
			this.BytesPanel.Controls.Add(this.byteEdit6);
			this.BytesPanel.Controls.Add(this.byteEdit5);
			this.BytesPanel.Controls.Add(this.byteEdit4);
			this.BytesPanel.Controls.Add(this.byteEdit3);
			this.BytesPanel.Controls.Add(this.byteEdit2);
			this.BytesPanel.Controls.Add(this.byteEdit1);
			this.BytesPanel.Location = new System.Drawing.Point(70, 20);
			this.BytesPanel.Name = "BytesPanel";
			this.BytesPanel.Size = new System.Drawing.Size(500, 250);
			this.BytesPanel.TabIndex = 89;
			// 
			// byteEdit16
			// 
			this.byteEdit16.Location = new System.Drawing.Point(0, 0);
			this.byteEdit16.Name = "byteEdit16";
			this.byteEdit16.Size = new System.Drawing.Size(25, 250);
			this.byteEdit16.TabIndex = 104;
			this.byteEdit16.Value = ((System.Byte)(76));
			this.byteEdit16.Visible = false;
			this.byteEdit16.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit15
			// 
			this.byteEdit15.Location = new System.Drawing.Point(30, 0);
			this.byteEdit15.Name = "byteEdit15";
			this.byteEdit15.Size = new System.Drawing.Size(25, 250);
			this.byteEdit15.TabIndex = 103;
			this.byteEdit15.Value = ((System.Byte)(101));
			this.byteEdit15.Visible = false;
			this.byteEdit15.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit14
			// 
			this.byteEdit14.Location = new System.Drawing.Point(60, 0);
			this.byteEdit14.Name = "byteEdit14";
			this.byteEdit14.Size = new System.Drawing.Size(25, 250);
			this.byteEdit14.TabIndex = 102;
			this.byteEdit14.Value = ((System.Byte)(119));
			this.byteEdit14.Visible = false;
			this.byteEdit14.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit13
			// 
			this.byteEdit13.Location = new System.Drawing.Point(90, 0);
			this.byteEdit13.Name = "byteEdit13";
			this.byteEdit13.Size = new System.Drawing.Size(25, 250);
			this.byteEdit13.TabIndex = 101;
			this.byteEdit13.Value = ((System.Byte)(105));
			this.byteEdit13.Visible = false;
			this.byteEdit13.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit12
			// 
			this.byteEdit12.Location = new System.Drawing.Point(120, 0);
			this.byteEdit12.Name = "byteEdit12";
			this.byteEdit12.Size = new System.Drawing.Size(25, 250);
			this.byteEdit12.TabIndex = 100;
			this.byteEdit12.Value = ((System.Byte)(115));
			this.byteEdit12.Visible = false;
			this.byteEdit12.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit11
			// 
			this.byteEdit11.Location = new System.Drawing.Point(150, 0);
			this.byteEdit11.Name = "byteEdit11";
			this.byteEdit11.Size = new System.Drawing.Size(25, 250);
			this.byteEdit11.TabIndex = 99;
			this.byteEdit11.Value = ((System.Byte)(77));
			this.byteEdit11.Visible = false;
			this.byteEdit11.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit10
			// 
			this.byteEdit10.Location = new System.Drawing.Point(180, 0);
			this.byteEdit10.Name = "byteEdit10";
			this.byteEdit10.Size = new System.Drawing.Size(25, 250);
			this.byteEdit10.TabIndex = 98;
			this.byteEdit10.Value = ((System.Byte)(111));
			this.byteEdit10.Visible = false;
			this.byteEdit10.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit9
			// 
			this.byteEdit9.Location = new System.Drawing.Point(210, 0);
			this.byteEdit9.Name = "byteEdit9";
			this.byteEdit9.Size = new System.Drawing.Size(25, 250);
			this.byteEdit9.TabIndex = 97;
			this.byteEdit9.Value = ((System.Byte)(116));
			this.byteEdit9.Visible = false;
			this.byteEdit9.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit8
			// 
			this.byteEdit8.Location = new System.Drawing.Point(260, 0);
			this.byteEdit8.Name = "byteEdit8";
			this.byteEdit8.Size = new System.Drawing.Size(25, 250);
			this.byteEdit8.TabIndex = 96;
			this.byteEdit8.Value = ((System.Byte)(101));
			this.byteEdit8.Visible = false;
			this.byteEdit8.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit7
			// 
			this.byteEdit7.Location = new System.Drawing.Point(290, 0);
			this.byteEdit7.Name = "byteEdit7";
			this.byteEdit7.Size = new System.Drawing.Size(25, 250);
			this.byteEdit7.TabIndex = 95;
			this.byteEdit7.Value = ((System.Byte)(110));
			this.byteEdit7.Visible = false;
			this.byteEdit7.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit6
			// 
			this.byteEdit6.Location = new System.Drawing.Point(320, 0);
			this.byteEdit6.Name = "byteEdit6";
			this.byteEdit6.Size = new System.Drawing.Size(25, 250);
			this.byteEdit6.TabIndex = 94;
			this.byteEdit6.Value = ((System.Byte)(46));
			this.byteEdit6.Visible = false;
			this.byteEdit6.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit5
			// 
			this.byteEdit5.Location = new System.Drawing.Point(350, 0);
			this.byteEdit5.Name = "byteEdit5";
			this.byteEdit5.Size = new System.Drawing.Size(25, 250);
			this.byteEdit5.TabIndex = 93;
			this.byteEdit5.Value = ((System.Byte)(99));
			this.byteEdit5.Visible = false;
			this.byteEdit5.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit4
			// 
			this.byteEdit4.Location = new System.Drawing.Point(380, 0);
			this.byteEdit4.Name = "byteEdit4";
			this.byteEdit4.Size = new System.Drawing.Size(25, 250);
			this.byteEdit4.TabIndex = 92;
			this.byteEdit4.Value = ((System.Byte)(111));
			this.byteEdit4.Visible = false;
			this.byteEdit4.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit3
			// 
			this.byteEdit3.Location = new System.Drawing.Point(410, 0);
			this.byteEdit3.Name = "byteEdit3";
			this.byteEdit3.Size = new System.Drawing.Size(25, 250);
			this.byteEdit3.TabIndex = 91;
			this.byteEdit3.Value = ((System.Byte)(109));
			this.byteEdit3.Visible = false;
			this.byteEdit3.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit2
			// 
			this.byteEdit2.Location = new System.Drawing.Point(440, 0);
			this.byteEdit2.Name = "byteEdit2";
			this.byteEdit2.Size = new System.Drawing.Size(25, 250);
			this.byteEdit2.TabIndex = 90;
			this.byteEdit2.Value = ((System.Byte)(0));
			this.byteEdit2.Visible = false;
			this.byteEdit2.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// byteEdit1
			// 
			this.byteEdit1.Location = new System.Drawing.Point(470, 0);
			this.byteEdit1.Name = "byteEdit1";
			this.byteEdit1.Size = new System.Drawing.Size(25, 250);
			this.byteEdit1.TabIndex = 89;
			this.byteEdit1.Value = ((System.Byte)(0));
			this.byteEdit1.Visible = false;
			this.byteEdit1.Changed += new System.EventHandler(this.byteEdit_Changed);
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(10, 110);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(23, 16);
			this.label5.TabIndex = 72;
			this.label5.Text = "Bits";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(10, 80);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(58, 16);
			this.label4.TabIndex = 71;
			this.label4.Text = "iso-8859-1";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(10, 50);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(45, 16);
			this.label3.TabIndex = 70;
			this.label3.Text = "Decimal";
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(10, 20);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(24, 16);
			this.label2.TabIndex = 69;
			this.label2.Text = "Hex";
			// 
			// vScrollBar1
			// 
			this.vScrollBar1.Dock = System.Windows.Forms.DockStyle.Right;
			this.vScrollBar1.Location = new System.Drawing.Point(565, 0);
			this.vScrollBar1.Name = "vScrollBar1";
			this.vScrollBar1.Size = new System.Drawing.Size(17, 115);
			this.vScrollBar1.TabIndex = 34;
			this.vScrollBar1.Resize += new System.EventHandler(this.vScrollBar1_Resize);
			this.vScrollBar1.ValueChanged += new System.EventHandler(this.vScrollBar1_ValueChanged);
			// 
			// listView1
			// 
			this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						this.columnHeader1,
																						this.columnHeader2,
																						this.columnHeader3});
			this.listView1.Dock = System.Windows.Forms.DockStyle.Fill;
			this.listView1.Font = new System.Drawing.Font("Courier New", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.listView1.FullRowSelect = true;
			this.listView1.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
																					  listViewItem1});
			this.listView1.Location = new System.Drawing.Point(0, 0);
			this.listView1.MultiSelect = false;
			this.listView1.Name = "listView1";
			this.listView1.Scrollable = false;
			this.listView1.Size = new System.Drawing.Size(565, 115);
			this.listView1.TabIndex = 36;
			this.listView1.View = System.Windows.Forms.View.Details;
			this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "Position";
			this.columnHeader1.Width = 70;
			// 
			// columnHeader2
			// 
			this.columnHeader2.Text = "Bytes";
			this.columnHeader2.Width = 343;
			// 
			// columnHeader3
			// 
			this.columnHeader3.Text = "Text";
			this.columnHeader3.Width = 126;
			// 
			// mainMenu1
			// 
			this.mainMenu1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.menuItem1});
			// 
			// menuItem1
			// 
			this.menuItem1.Index = 0;
			this.menuItem1.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																					  this.OpenFileMenuItem,
																					  this.AuthorsSiteMenuItem,
																					  this.ExitProgram});
			this.menuItem1.Text = "File";
			// 
			// OpenFileMenuItem
			// 
			this.OpenFileMenuItem.Index = 0;
			this.OpenFileMenuItem.Text = "Open ...";
			this.OpenFileMenuItem.Click += new System.EventHandler(this.OpenFileMenuItem_Click);
			// 
			// AuthorsSiteMenuItem
			// 
			this.AuthorsSiteMenuItem.Index = 1;
			this.AuthorsSiteMenuItem.Text = "www.lewismoten.com";
			this.AuthorsSiteMenuItem.Click += new System.EventHandler(this.AuthorsSiteMenuItem_Click);
			// 
			// ExitProgram
			// 
			this.ExitProgram.Index = 2;
			this.ExitProgram.Text = "Exit";
			this.ExitProgram.Click += new System.EventHandler(this.ExitProgramMenuItem_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(582, 395);
			this.Controls.Add(this.listView1);
			this.Controls.Add(this.vScrollBar1);
			this.Controls.Add(this.BytesGroup);
			this.MaximumSize = new System.Drawing.Size(590, 2000);
			this.Menu = this.mainMenu1;
			this.MinimumSize = new System.Drawing.Size(590, 450);
			this.Name = "Form1";
			this.Text = "Mini Hex Edit";
			this.Load += new System.EventHandler(this.Form1_Load);
			this.BytesGroup.ResumeLayout(false);
			this.BytesPanel.ResumeLayout(false);
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

		private string _fileName;
		private long _length;
		private int VisibleItems
		{
			get
			{
				int itemHeight = 15;
				int headerFooter = 22;
				int items = (int)Math.Ceiling((listView1.Height - headerFooter) / itemHeight);
				if(items <= 1) items = 1;
				return items;
			}
		}
		private string FileName
		{
			get
			{
				return _fileName;
			}
			set
			{
				_fileName = value;
				Text = System.IO.Path.GetFileName(_fileName) + " - Mini Hex Edit";
				_length = new System.IO.FileInfo(_fileName).Length;
				vScrollBar1.Value = 0;
				vScrollBar1.Maximum = (int)Math.Floor(_length / 16);
				vScrollBar1.LargeChange = VisibleItems;
			}
		}
		private long Length
		{
			get
			{
				return _length;
			}
		}
		private void OpenFileMenuItem_Click(object sender, System.EventArgs e)
		{
			if(openFileDialog1.ShowDialog(this) == DialogResult.Cancel) return;
			FileName = openFileDialog1.FileName;
			SelectedIndex = 0;
			SetFirstLine(0);
			UpdateBytes.Visible = true;
		}
		private void SetFirstLine(int lineIndex)
		{
			listView1.Items.Clear();
			if(FileName == null) return;
			System.IO.Stream stream = null;
			try
			{
				stream = System.IO.File.OpenRead(FileName);
				for(int i = 0; i < VisibleItems; i++)
				{
					int index = lineIndex + (i * 16);
					if(index >= Length) return;
					stream.Seek(index, System.IO.SeekOrigin.Begin);
					byte[] buffer = new byte[16];
					if(buffer.Length > stream.Length - index)
						buffer = new byte[stream.Length - index];
					stream.Read(buffer, 0, buffer.Length);
					addLine(index, buffer);
				}
			}
			finally
			{
				if(stream != null) stream.Close();
			}
		}
		private void addLine(long index, byte[] buffer)
		{
			ListViewItem item = new ListViewItem();
			item.Text = index.ToString("x").PadLeft(8, '0') + "h";
			item.SubItems.Add(BitConverter.ToString(buffer));
			StringBuilder sb = new StringBuilder();
			for(int i = 0; i < buffer.Length; i++)
				if(buffer[i] <= 31 || (buffer[i] >= 127 && buffer[i] <= 159)) sb.Append(".");
				else sb.Append(Encoder().GetString(buffer, i, 1));
			item.SubItems.Add(sb.ToString());
			item.Selected = SelectedIndex == vScrollBar1.Value + listView1.Items.Count;
			listView1.Items.Add(item);
		}
		Encoding _encoder = null;
		protected Encoding Encoder()
		{
			if(_encoder == null) _encoder = Encoding.GetEncoding("iso-8859-1");
			return _encoder;
		}

		private void vScrollBar1_ValueChanged(object sender, System.EventArgs e)
		{
			SetFirstLine(vScrollBar1.Value * 16);
		}

		int _selectedIndex = 0;
		protected int SelectedIndex
		{
			get
			{
				return _selectedIndex;
			}
			set
			{
				_selectedIndex = value;
				ShowBytes();
			}
		}
		private void listView1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if(listView1.SelectedItems.Count == 0) return;
			SelectedIndex = this.listView1.SelectedItems[0].Index + this.vScrollBar1.Value;
		}

		private void vScrollBar1_Resize(object sender, System.EventArgs e)
		{
			vScrollBar1.LargeChange = VisibleItems;
			SetFirstLine(vScrollBar1.Value * 16);
		}
		private void ShowBytes()
		{
			if(FileName == null) return;
			int index = SelectedIndex * 16;
			if(index >= Length) return;

			byte[] buffer = new byte[16];
			System.IO.Stream stream = null;
			try
			{
				stream = System.IO.File.OpenRead(FileName);
				stream.Seek(SelectedIndex * 16, System.IO.SeekOrigin.Begin);

				if(buffer.Length > stream.Length - index)
					buffer = new byte[stream.Length - index];

				stream.Read(buffer, 0, buffer.Length);
			}
			finally
			{
				if(stream != null) stream.Close();
			}

			for(int i = 0; i < 16; i++)
				if(i < buffer.Length)
				{
					((ByteEdit)BytesPanel.Controls[i]).Value = buffer[i];
					BytesPanel.Controls[i].Visible = true;
				}
				else
					BytesPanel.Controls[i].Visible = false;

			BytesGroup.Text = 
				"Bytes " 
				+ (SelectedIndex * 16).ToString("x").PadLeft(8, '0') 
				+ "h - " 
				+ ((SelectedIndex * 16) + (buffer.Length - 1)).ToString("x").PadLeft(8, '0') 
				+ "h ("
				+ (SelectedIndex * 16).ToString() 
				+ " - "
				+ ((SelectedIndex * 16) + (buffer.Length - 1)).ToString()
				+ ")";

			UpdateBytes.Enabled = false;
		}

		private void UpdateBytes_Click(object sender, System.EventArgs e)
		{
			if(FileName == null) return;
			int index = SelectedIndex * 16;
			if(index >= Length) return;

			byte[] buffer = new byte[16];
			if(buffer.Length > Length - index)
				buffer = new byte[Length - index];

			for(int i = 0; i < buffer.Length; i++)
				buffer[i] = ((ByteEdit)BytesPanel.Controls[i]).Value;


			System.IO.Stream stream = null;
			try
			{
				stream = System.IO.File.OpenWrite(FileName);
				stream.Seek(SelectedIndex * 16, System.IO.SeekOrigin.Begin);
				stream.Write(buffer, 0, buffer.Length);
			}
			finally
			{
				if(stream != null) stream.Close();
			}
			SetFirstLine(vScrollBar1.Value * 16);
			UpdateBytes.Enabled = false;
		}

		private void byteEdit_Changed(object sender, System.EventArgs e)
		{
			UpdateBytes.Enabled = true;
		}

		private void AuthorsSiteMenuItem_Click(object sender, System.EventArgs e)
		{
			System.Diagnostics.Process.Start("http://www.lewismoten.com");
		}

		private void ExitProgramMenuItem_Click(object sender, System.EventArgs e)
		{
			Application.Exit();
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			MaximumSize = new Size(MaximumSize.Width, Screen.PrimaryScreen.Bounds.Height);

			Stream stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream("MiniHexEdit.App.ico"); 
			System.Drawing.Icon icon = new System.Drawing.Icon(stream); 
			this.Icon = icon;

			Splash splash = new Splash();
			splash.Icon = icon;
			splash.ShowDialog(this);
			splash.Dispose();
			splash = null;
		}

	}
}
