using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace EmbeddedText
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class DemonstrationForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button ShowConstitution;
		private System.Windows.Forms.Button ShowBillOfRights;
		private System.Windows.Forms.TextBox Textarea;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DemonstrationForm()
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
			this.ShowConstitution = new System.Windows.Forms.Button();
			this.ShowBillOfRights = new System.Windows.Forms.Button();
			this.Textarea = new System.Windows.Forms.TextBox();
			this.SuspendLayout();
			// 
			// ShowConstitution
			// 
			this.ShowConstitution.Location = new System.Drawing.Point(30, 10);
			this.ShowConstitution.Name = "ShowConstitution";
			this.ShowConstitution.Size = new System.Drawing.Size(110, 23);
			this.ShowConstitution.TabIndex = 0;
			this.ShowConstitution.Text = "Show Constitution";
			this.ShowConstitution.Click += new System.EventHandler(this.ShowConstitution_Click);
			// 
			// ShowBillOfRights
			// 
			this.ShowBillOfRights.Location = new System.Drawing.Point(210, 10);
			this.ShowBillOfRights.Name = "ShowBillOfRights";
			this.ShowBillOfRights.Size = new System.Drawing.Size(110, 23);
			this.ShowBillOfRights.TabIndex = 1;
			this.ShowBillOfRights.Text = "Show Bill of Rights";
			this.ShowBillOfRights.Click += new System.EventHandler(this.ShowBillOfRights_Click);
			// 
			// Textarea
			// 
			this.Textarea.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.Textarea.Location = new System.Drawing.Point(30, 40);
			this.Textarea.Multiline = true;
			this.Textarea.Name = "Textarea";
			this.Textarea.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.Textarea.Size = new System.Drawing.Size(290, 200);
			this.Textarea.TabIndex = 2;
			this.Textarea.Text = "";
			// 
			// DemonstrationForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(342, 266);
			this.Controls.Add(this.Textarea);
			this.Controls.Add(this.ShowBillOfRights);
			this.Controls.Add(this.ShowConstitution);
			this.Name = "DemonstrationForm";
			this.Text = "Embedded Text Demo";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new DemonstrationForm());
		}

		private void ShowConstitution_Click(object sender, System.EventArgs e)
		{
			Textarea.Text = Resources.GetEmbeddedText("constitution.txt");
		}

		private void ShowBillOfRights_Click(object sender, System.EventArgs e)
		{
			Textarea.Text = Resources.GetEmbeddedText("Folder.BillOfRights.txt");
		}
	}
}
