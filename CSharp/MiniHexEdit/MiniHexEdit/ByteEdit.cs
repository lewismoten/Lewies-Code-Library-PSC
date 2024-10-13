using System;
using System.Collections;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Windows.Forms;
using System.Text;

namespace MiniHexEdit
{
	/// <summary>
	/// Summary description for ByteEdit.
	/// </summary>
	public class ByteEdit : System.Windows.Forms.UserControl
	{
		/// <summary> 
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;
		public event EventHandler Changed;

		public ByteEdit()
		{
			// This call is required by the Windows.Forms Form Designer.
			InitializeComponent();

			// TODO: Add any initialization after the InitializeComponent call

		}
		private System.Windows.Forms.TextBox TextValue;
		private System.Windows.Forms.TextBox HexValue;
		private System.Windows.Forms.TextBox DecimalValue;
		byte _value = 0;
		private System.Windows.Forms.Panel BitPanel;
		private System.Windows.Forms.CheckBox Bit1;
		private System.Windows.Forms.CheckBox Bit2;
		private System.Windows.Forms.CheckBox Bit3;
		private System.Windows.Forms.CheckBox Bit4;
		private System.Windows.Forms.CheckBox Bit5;
		private System.Windows.Forms.CheckBox Bit6;
		private System.Windows.Forms.CheckBox Bit7;
		private System.Windows.Forms.CheckBox Bit8;
	
		private bool updating = false;
		public byte Value
		{
			get
			{
				return _value;
			}
			set
			{
				if(updating) return;
				if(value == _value) return;
				updating = true;
				_value = value;
				HexValue.Text = _value.ToString("x").PadLeft(2, '0');
				DecimalValue.Text = _value.ToString();
				if(_value <= 31 || (_value >= 127 && _value <= 159)) 
				{
					TextValue.Text = "";
					TextValue.BackColor = SystemColors.Control;
					TextValue.ForeColor = SystemColors.ControlText;
				}
				else 
				{
					TextValue.Text = Encoder().GetString(new byte[]{_value}, 0, 1);
					TextValue.BackColor = SystemColors.Window;
					TextValue.ForeColor = SystemColors.WindowText;
				}
				BitArray bits = new BitArray(new byte[]{_value});
				for(int i = 0; i < 8; i++)
					((CheckBox)BitPanel.Controls[i]).Checked = bits[i];
				updating = false;
				try
				{
					Changed(this, new EventArgs());
				}
				catch(Exception)
				{
				}
			}
		}
		Encoding _encoder = null;
		protected Encoding Encoder()
		{
			if(_encoder == null) _encoder = Encoding.GetEncoding("iso-8859-1");
			return _encoder;
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

		#region Component Designer generated code
		/// <summary> 
		/// Required method for Designer support - do not modify 
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			this.TextValue = new System.Windows.Forms.TextBox();
			this.HexValue = new System.Windows.Forms.TextBox();
			this.DecimalValue = new System.Windows.Forms.TextBox();
			this.BitPanel = new System.Windows.Forms.Panel();
			this.Bit1 = new System.Windows.Forms.CheckBox();
			this.Bit2 = new System.Windows.Forms.CheckBox();
			this.Bit3 = new System.Windows.Forms.CheckBox();
			this.Bit4 = new System.Windows.Forms.CheckBox();
			this.Bit5 = new System.Windows.Forms.CheckBox();
			this.Bit6 = new System.Windows.Forms.CheckBox();
			this.Bit7 = new System.Windows.Forms.CheckBox();
			this.Bit8 = new System.Windows.Forms.CheckBox();
			this.BitPanel.SuspendLayout();
			this.SuspendLayout();
			// 
			// TextValue
			// 
			this.TextValue.Location = new System.Drawing.Point(0, 60);
			this.TextValue.MaxLength = 1;
			this.TextValue.Name = "TextValue";
			this.TextValue.Size = new System.Drawing.Size(25, 20);
			this.TextValue.TabIndex = 2;
			this.TextValue.Text = "";
			this.TextValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
			this.TextValue.TextChanged += new System.EventHandler(this.TextValue_TextChanged);
			// 
			// HexValue
			// 
			this.HexValue.Location = new System.Drawing.Point(0, 0);
			this.HexValue.MaxLength = 2;
			this.HexValue.Name = "HexValue";
			this.HexValue.Size = new System.Drawing.Size(25, 20);
			this.HexValue.TabIndex = 0;
			this.HexValue.Text = "00";
			this.HexValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.HexValue.Validating += new System.ComponentModel.CancelEventHandler(this.HexValue_Validating);
			this.HexValue.Validated += new System.EventHandler(this.HexValue_Validated);
			// 
			// DecimalValue
			// 
			this.DecimalValue.Location = new System.Drawing.Point(0, 30);
			this.DecimalValue.MaxLength = 3;
			this.DecimalValue.Name = "DecimalValue";
			this.DecimalValue.Size = new System.Drawing.Size(25, 20);
			this.DecimalValue.TabIndex = 1;
			this.DecimalValue.Text = "0";
			this.DecimalValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
			this.DecimalValue.Validating += new System.ComponentModel.CancelEventHandler(this.DecimalValue_Validating);
			this.DecimalValue.Validated += new System.EventHandler(this.DecimalValue_Validated);
			// 
			// BitPanel
			// 
			this.BitPanel.Controls.Add(this.Bit8);
			this.BitPanel.Controls.Add(this.Bit7);
			this.BitPanel.Controls.Add(this.Bit6);
			this.BitPanel.Controls.Add(this.Bit5);
			this.BitPanel.Controls.Add(this.Bit4);
			this.BitPanel.Controls.Add(this.Bit3);
			this.BitPanel.Controls.Add(this.Bit2);
			this.BitPanel.Controls.Add(this.Bit1);
			this.BitPanel.Location = new System.Drawing.Point(5, 90);
			this.BitPanel.Name = "BitPanel";
			this.BitPanel.Size = new System.Drawing.Size(15, 155);
			this.BitPanel.TabIndex = 11;
			// 
			// Bit1
			// 
			this.Bit1.Location = new System.Drawing.Point(0, 0);
			this.Bit1.Name = "Bit1";
			this.Bit1.Size = new System.Drawing.Size(15, 15);
			this.Bit1.TabIndex = 11;
			this.Bit1.TabStop = false;
			this.Bit1.CheckedChanged += new System.EventHandler(this.Bit_CheckedChanged);
			// 
			// Bit2
			// 
			this.Bit2.Location = new System.Drawing.Point(0, 20);
			this.Bit2.Name = "Bit2";
			this.Bit2.Size = new System.Drawing.Size(15, 15);
			this.Bit2.TabIndex = 12;
			this.Bit2.TabStop = false;
			this.Bit2.CheckedChanged += new System.EventHandler(this.Bit_CheckedChanged);
			// 
			// Bit3
			// 
			this.Bit3.Location = new System.Drawing.Point(0, 40);
			this.Bit3.Name = "Bit3";
			this.Bit3.Size = new System.Drawing.Size(15, 15);
			this.Bit3.TabIndex = 13;
			this.Bit3.TabStop = false;
			this.Bit3.CheckedChanged += new System.EventHandler(this.Bit_CheckedChanged);
			// 
			// Bit4
			// 
			this.Bit4.Location = new System.Drawing.Point(0, 60);
			this.Bit4.Name = "Bit4";
			this.Bit4.Size = new System.Drawing.Size(15, 15);
			this.Bit4.TabIndex = 14;
			this.Bit4.TabStop = false;
			this.Bit4.CheckedChanged += new System.EventHandler(this.Bit_CheckedChanged);
			// 
			// Bit5
			// 
			this.Bit5.Location = new System.Drawing.Point(0, 80);
			this.Bit5.Name = "Bit5";
			this.Bit5.Size = new System.Drawing.Size(15, 15);
			this.Bit5.TabIndex = 15;
			this.Bit5.TabStop = false;
			this.Bit5.CheckedChanged += new System.EventHandler(this.Bit_CheckedChanged);
			// 
			// Bit6
			// 
			this.Bit6.Location = new System.Drawing.Point(0, 100);
			this.Bit6.Name = "Bit6";
			this.Bit6.Size = new System.Drawing.Size(15, 15);
			this.Bit6.TabIndex = 16;
			this.Bit6.TabStop = false;
			this.Bit6.CheckedChanged += new System.EventHandler(this.Bit_CheckedChanged);
			// 
			// Bit7
			// 
			this.Bit7.Location = new System.Drawing.Point(0, 120);
			this.Bit7.Name = "Bit7";
			this.Bit7.Size = new System.Drawing.Size(15, 15);
			this.Bit7.TabIndex = 17;
			this.Bit7.TabStop = false;
			this.Bit7.CheckedChanged += new System.EventHandler(this.Bit_CheckedChanged);
			// 
			// Bit8
			// 
			this.Bit8.Location = new System.Drawing.Point(0, 140);
			this.Bit8.Name = "Bit8";
			this.Bit8.Size = new System.Drawing.Size(15, 15);
			this.Bit8.TabIndex = 18;
			this.Bit8.TabStop = false;
			this.Bit8.CheckedChanged += new System.EventHandler(this.Bit_CheckedChanged);
			// 
			// ByteEdit
			// 
			this.Controls.Add(this.BitPanel);
			this.Controls.Add(this.TextValue);
			this.Controls.Add(this.HexValue);
			this.Controls.Add(this.DecimalValue);
			this.Name = "ByteEdit";
			this.Size = new System.Drawing.Size(25, 245);
			this.BitPanel.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion

		private void Bit_CheckedChanged(object sender, System.EventArgs e)
		{
			byte[] buffer = new byte[1];
			BitArray bits = new BitArray(8, false);
			for(int i = 0; i < 8; i++)
				bits[i] = ((CheckBox)BitPanel.Controls[i]).Checked;
			bits.CopyTo(buffer, 0);
			Value = buffer[0];
		}

		private void HexValue_Validated(object sender, System.EventArgs e)
		{
			Value = byte.Parse(HexValue.Text, System.Globalization.NumberStyles.HexNumber);
		}

		private void DecimalValue_Validated(object sender, System.EventArgs e)
		{
			Value = byte.Parse(DecimalValue.Text, System.Globalization.NumberStyles.Number);
		}

		private void HexValue_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string hex = HexValue.Text;
			for(int i = 0; i < hex.Length; i++)
				if("0123456789abcdef".IndexOf(hex.Substring(0, 1).ToLower()) == -1)
				{
					e.Cancel = true; 
					MessageBox.Show("An integer between 00 and FF is required.");
					return;
				}
		}

		private void DecimalValue_Validating(object sender, System.ComponentModel.CancelEventArgs e)
		{
			string number = DecimalValue.Text;
			for(int i = 0; i < number.Length; i++)
				if("0123456789".IndexOf(number.Substring(i, 1)) == -1)
				{
					e.Cancel = true;
					MessageBox.Show("An integer between 0 and 255 is required.");
					return;
				}
			if(int.Parse(number) > 255)
			{
				e.Cancel = true;
				MessageBox.Show("An integer between 0 and 255 is required.");
				return;
			}
		}

		private void TextValue_TextChanged(object sender, System.EventArgs e)
		{
			if(TextValue.Text.Length == 0)
				Value = 0;
			else
				Value = Encoder().GetBytes(TextValue.Text)[0];
		}
	}
}
