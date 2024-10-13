using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Web.Mail;

namespace SendMail
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.TextBox ServerText;
		private System.Windows.Forms.TextBox UsernameText;
		private System.Windows.Forms.TextBox PasswordText;
		private System.Windows.Forms.TextBox FromText;
		private System.Windows.Forms.TextBox ToText;
		private System.Windows.Forms.TextBox SubjectText;
		private System.Windows.Forms.TextBox MessageText;
		private System.Windows.Forms.Button SendButton;
		private System.Windows.Forms.Label ServerLabel;
		private System.Windows.Forms.Label UsernameLabel;
		private System.Windows.Forms.Label PasswordLabel;
		private System.Windows.Forms.Label FromLabel;
		private System.Windows.Forms.Label ToLabel;
		private System.Windows.Forms.Label SubjectLabel;
		private System.Windows.Forms.Label MessageLabel;
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
			this.ServerLabel = new System.Windows.Forms.Label();
			this.UsernameLabel = new System.Windows.Forms.Label();
			this.PasswordLabel = new System.Windows.Forms.Label();
			this.FromLabel = new System.Windows.Forms.Label();
			this.ToLabel = new System.Windows.Forms.Label();
			this.SubjectLabel = new System.Windows.Forms.Label();
			this.MessageLabel = new System.Windows.Forms.Label();
			this.ServerText = new System.Windows.Forms.TextBox();
			this.UsernameText = new System.Windows.Forms.TextBox();
			this.PasswordText = new System.Windows.Forms.TextBox();
			this.FromText = new System.Windows.Forms.TextBox();
			this.ToText = new System.Windows.Forms.TextBox();
			this.SubjectText = new System.Windows.Forms.TextBox();
			this.MessageText = new System.Windows.Forms.TextBox();
			this.SendButton = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// ServerLabel
			// 
			this.ServerLabel.AutoSize = true;
			this.ServerLabel.Location = new System.Drawing.Point(10, 10);
			this.ServerLabel.Name = "ServerLabel";
			this.ServerLabel.Size = new System.Drawing.Size(38, 16);
			this.ServerLabel.TabIndex = 0;
			this.ServerLabel.Text = "Server";
			// 
			// UsernameLabel
			// 
			this.UsernameLabel.AutoSize = true;
			this.UsernameLabel.Location = new System.Drawing.Point(10, 40);
			this.UsernameLabel.Name = "UsernameLabel";
			this.UsernameLabel.Size = new System.Drawing.Size(56, 16);
			this.UsernameLabel.TabIndex = 1;
			this.UsernameLabel.Text = "Username";
			// 
			// PasswordLabel
			// 
			this.PasswordLabel.AutoSize = true;
			this.PasswordLabel.Location = new System.Drawing.Point(190, 40);
			this.PasswordLabel.Name = "PasswordLabel";
			this.PasswordLabel.Size = new System.Drawing.Size(54, 16);
			this.PasswordLabel.TabIndex = 2;
			this.PasswordLabel.Text = "Password";
			// 
			// FromLabel
			// 
			this.FromLabel.AutoSize = true;
			this.FromLabel.Location = new System.Drawing.Point(10, 90);
			this.FromLabel.Name = "FromLabel";
			this.FromLabel.Size = new System.Drawing.Size(75, 16);
			this.FromLabel.TabIndex = 3;
			this.FromLabel.Text = "From Address";
			// 
			// ToLabel
			// 
			this.ToLabel.AutoSize = true;
			this.ToLabel.Location = new System.Drawing.Point(190, 90);
			this.ToLabel.Name = "ToLabel";
			this.ToLabel.Size = new System.Drawing.Size(62, 16);
			this.ToLabel.TabIndex = 5;
			this.ToLabel.Text = "To Address";
			// 
			// SubjectLabel
			// 
			this.SubjectLabel.AutoSize = true;
			this.SubjectLabel.Location = new System.Drawing.Point(10, 150);
			this.SubjectLabel.Name = "SubjectLabel";
			this.SubjectLabel.Size = new System.Drawing.Size(42, 16);
			this.SubjectLabel.TabIndex = 6;
			this.SubjectLabel.Text = "Subject";
			// 
			// MessageLabel
			// 
			this.MessageLabel.AutoSize = true;
			this.MessageLabel.Location = new System.Drawing.Point(10, 190);
			this.MessageLabel.Name = "MessageLabel";
			this.MessageLabel.Size = new System.Drawing.Size(50, 16);
			this.MessageLabel.TabIndex = 7;
			this.MessageLabel.Text = "Message";
			// 
			// ServerText
			// 
			this.ServerText.Location = new System.Drawing.Point(70, 10);
			this.ServerText.Name = "ServerText";
			this.ServerText.Size = new System.Drawing.Size(280, 20);
			this.ServerText.TabIndex = 8;
			this.ServerText.Text = "";
			// 
			// UsernameText
			// 
			this.UsernameText.Location = new System.Drawing.Point(10, 60);
			this.UsernameText.Name = "UsernameText";
			this.UsernameText.Size = new System.Drawing.Size(160, 20);
			this.UsernameText.TabIndex = 9;
			this.UsernameText.Text = "";
			// 
			// PasswordText
			// 
			this.PasswordText.Location = new System.Drawing.Point(190, 60);
			this.PasswordText.Name = "PasswordText";
			this.PasswordText.PasswordChar = '*';
			this.PasswordText.Size = new System.Drawing.Size(160, 20);
			this.PasswordText.TabIndex = 10;
			this.PasswordText.Text = "";
			// 
			// FromText
			// 
			this.FromText.Location = new System.Drawing.Point(10, 110);
			this.FromText.Name = "FromText";
			this.FromText.Size = new System.Drawing.Size(160, 20);
			this.FromText.TabIndex = 11;
			this.FromText.Text = "";
			// 
			// ToText
			// 
			this.ToText.Location = new System.Drawing.Point(190, 110);
			this.ToText.Name = "ToText";
			this.ToText.Size = new System.Drawing.Size(160, 20);
			this.ToText.TabIndex = 12;
			this.ToText.Text = "";
			// 
			// SubjectText
			// 
			this.SubjectText.Location = new System.Drawing.Point(70, 150);
			this.SubjectText.Name = "SubjectText";
			this.SubjectText.Size = new System.Drawing.Size(280, 20);
			this.SubjectText.TabIndex = 13;
			this.SubjectText.Text = "";
			// 
			// MessageText
			// 
			this.MessageText.Location = new System.Drawing.Point(70, 180);
			this.MessageText.Multiline = true;
			this.MessageText.Name = "MessageText";
			this.MessageText.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.MessageText.Size = new System.Drawing.Size(280, 150);
			this.MessageText.TabIndex = 14;
			this.MessageText.Text = "";
			// 
			// SendButton
			// 
			this.SendButton.Location = new System.Drawing.Point(280, 340);
			this.SendButton.Name = "SendButton";
			this.SendButton.TabIndex = 15;
			this.SendButton.Text = "Send";
			this.SendButton.Click += new System.EventHandler(this.SendButton_Click);
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(362, 376);
			this.Controls.Add(this.SendButton);
			this.Controls.Add(this.MessageText);
			this.Controls.Add(this.SubjectText);
			this.Controls.Add(this.ToText);
			this.Controls.Add(this.FromText);
			this.Controls.Add(this.PasswordText);
			this.Controls.Add(this.UsernameText);
			this.Controls.Add(this.ServerText);
			this.Controls.Add(this.MessageLabel);
			this.Controls.Add(this.SubjectLabel);
			this.Controls.Add(this.ToLabel);
			this.Controls.Add(this.FromLabel);
			this.Controls.Add(this.PasswordLabel);
			this.Controls.Add(this.UsernameLabel);
			this.Controls.Add(this.ServerLabel);
			this.Name = "Form1";
			this.Text = "Relay mail through CDO";
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

		private void SendButton_Click(object sender, System.EventArgs e)
		{
			MailMessage message = new MailMessage();
			message.From = FromText.Text;
			message.To = ToText.Text;
			message.Subject = SubjectText.Text;
			message.Body = MessageText.Text;
			
			Smtp.Send(message, ServerText.Text, UsernameText.Text, PasswordText.Text);

			ToText.Text = "";
			SubjectText.Text = "";
			MessageText.Text = "";
		}

	}
}
