using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Security.Cryptography;

namespace rsa
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button Encrypt;
		private System.Windows.Forms.TextBox txtPublic;
		private System.Windows.Forms.TextBox txtPrivate;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.TextBox txtClear;
		private System.Windows.Forms.TextBox txtEncrypted;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Button Decrypt;
		private System.Windows.Forms.Button NewKeys;
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
			this.Encrypt = new System.Windows.Forms.Button();
			this.txtPublic = new System.Windows.Forms.TextBox();
			this.txtPrivate = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.txtClear = new System.Windows.Forms.TextBox();
			this.txtEncrypted = new System.Windows.Forms.TextBox();
			this.label4 = new System.Windows.Forms.Label();
			this.Decrypt = new System.Windows.Forms.Button();
			this.NewKeys = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// Encrypt
			// 
			this.Encrypt.Location = new System.Drawing.Point(608, 8);
			this.Encrypt.Name = "Encrypt";
			this.Encrypt.TabIndex = 0;
			this.Encrypt.Text = "Encrypt";
			this.Encrypt.Click += new System.EventHandler(this.Encrypt_Click);
			// 
			// txtPublic
			// 
			this.txtPublic.Location = new System.Drawing.Point(8, 40);
			this.txtPublic.Multiline = true;
			this.txtPublic.Name = "txtPublic";
			this.txtPublic.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.txtPublic.Size = new System.Drawing.Size(408, 72);
			this.txtPublic.TabIndex = 1;
			this.txtPublic.Text = "";
			// 
			// txtPrivate
			// 
			this.txtPrivate.Location = new System.Drawing.Point(8, 144);
			this.txtPrivate.Multiline = true;
			this.txtPrivate.Name = "txtPrivate";
			this.txtPrivate.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.txtPrivate.Size = new System.Drawing.Size(408, 232);
			this.txtPrivate.TabIndex = 2;
			this.txtPrivate.Text = "";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.TabIndex = 3;
			this.label1.Text = "Public Key";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 120);
			this.label2.Name = "label2";
			this.label2.TabIndex = 4;
			this.label2.Text = "Private Key";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(432, 16);
			this.label3.Name = "label3";
			this.label3.TabIndex = 5;
			this.label3.Text = "Message";
			// 
			// txtClear
			// 
			this.txtClear.Location = new System.Drawing.Point(432, 40);
			this.txtClear.Multiline = true;
			this.txtClear.Name = "txtClear";
			this.txtClear.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.txtClear.Size = new System.Drawing.Size(248, 72);
			this.txtClear.TabIndex = 6;
			this.txtClear.Text = "This is a secret message.";
			// 
			// txtEncrypted
			// 
			this.txtEncrypted.Location = new System.Drawing.Point(432, 144);
			this.txtEncrypted.Multiline = true;
			this.txtEncrypted.Name = "txtEncrypted";
			this.txtEncrypted.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.txtEncrypted.Size = new System.Drawing.Size(248, 224);
			this.txtEncrypted.TabIndex = 7;
			this.txtEncrypted.Text = "";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(432, 120);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(160, 23);
			this.label4.TabIndex = 8;
			this.label4.Text = "Encrypted Message";
			// 
			// Decrypt
			// 
			this.Decrypt.Location = new System.Drawing.Point(600, 112);
			this.Decrypt.Name = "Decrypt";
			this.Decrypt.TabIndex = 9;
			this.Decrypt.Text = "Decrypt";
			this.Decrypt.Click += new System.EventHandler(this.Decrypt_Click);
			// 
			// NewKeys
			// 
			this.NewKeys.Location = new System.Drawing.Point(336, 8);
			this.NewKeys.Name = "NewKeys";
			this.NewKeys.TabIndex = 10;
			this.NewKeys.Text = "New Keys";
			this.NewKeys.Click += new System.EventHandler(this.NewKeys_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(704, 382);
			this.Controls.Add(this.NewKeys);
			this.Controls.Add(this.Decrypt);
			this.Controls.Add(this.txtEncrypted);
			this.Controls.Add(this.txtClear);
			this.Controls.Add(this.txtPrivate);
			this.Controls.Add(this.txtPublic);
			this.Controls.Add(this.Encrypt);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label4);
			this.Name = "Form1";
			this.Text = "RSA Example";
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

		private void NewKeys_Click(object sender, System.EventArgs e)
		{
			try
			{
				CspParameters cspParam = new CspParameters();
				cspParam.Flags = CspProviderFlags.UseMachineKeyStore;
				System.Security.Cryptography.RSACryptoServiceProvider RSA = new System.Security.Cryptography.RSACryptoServiceProvider(cspParam);
				txtPublic.Text = RSA.ToXmlString(false);
				txtPrivate.Text = RSA.ToXmlString(true);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void Encrypt_Click(object sender, System.EventArgs e)
		{
			try
			{
				RSACryptoServiceProvider RSA = new RSACryptoServiceProvider();
				RSA.FromXmlString(txtPublic.Text);
				byte[] decrypted = System.Text.Encoding.Unicode.GetBytes(txtClear.Text);
				byte[] encrypted = RSA.Encrypt(decrypted, false);
				txtEncrypted.Text = System.Convert.ToBase64String(encrypted);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}
		}

		private void Decrypt_Click(object sender, System.EventArgs e)
		{
			try
			{
				CspParameters cspParam = new CspParameters();
				cspParam.Flags = CspProviderFlags.UseMachineKeyStore;
				RSACryptoServiceProvider RSA = new RSACryptoServiceProvider(cspParam);
				RSA.FromXmlString(txtPrivate.Text);
				byte[] encrypted = System.Convert.FromBase64String(txtEncrypted.Text);
				byte[] decrypted = RSA.Decrypt(encrypted, false);
				txtClear.Text = System.Text.Encoding.Unicode.GetString(decrypted);
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.Message);
			}

		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			NewKeys_Click(sender, e);
			Encrypt_Click(sender, e);
		}

	}
}
