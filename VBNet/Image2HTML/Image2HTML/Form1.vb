Public Class Form1
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
	Friend WithEvents Panel1 As System.Windows.Forms.Panel
	Friend WithEvents btnBrowse As System.Windows.Forms.Button
	Friend WithEvents btnOK As System.Windows.Forms.Button
	Friend WithEvents Panel2 As System.Windows.Forms.Panel
	Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
	Friend WithEvents FileName As System.Windows.Forms.TextBox
	Friend WithEvents Picture As System.Windows.Forms.PictureBox
	Friend WithEvents ScriptInversion As System.Windows.Forms.CheckBox
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
		Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
		Me.Panel1 = New System.Windows.Forms.Panel
		Me.FileName = New System.Windows.Forms.TextBox
		Me.Panel2 = New System.Windows.Forms.Panel
		Me.btnOK = New System.Windows.Forms.Button
		Me.btnBrowse = New System.Windows.Forms.Button
		Me.Picture = New System.Windows.Forms.PictureBox
		Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
		Me.ScriptInversion = New System.Windows.Forms.CheckBox
		Me.Panel1.SuspendLayout()
		Me.Panel2.SuspendLayout()
		Me.SuspendLayout()
		'
		'Panel1
		'
		Me.Panel1.Controls.Add(Me.FileName)
		Me.Panel1.Controls.Add(Me.Panel2)
		Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.Panel1.Location = New System.Drawing.Point(0, 236)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Size = New System.Drawing.Size(462, 30)
		Me.Panel1.TabIndex = 0
		'
		'FileName
		'
		Me.FileName.Dock = System.Windows.Forms.DockStyle.Top
		Me.FileName.Location = New System.Drawing.Point(0, 0)
		Me.FileName.Name = "FileName"
		Me.FileName.Size = New System.Drawing.Size(232, 20)
		Me.FileName.TabIndex = 4
		Me.FileName.Text = "TextBox1"
		'
		'Panel2
		'
		Me.Panel2.Controls.Add(Me.btnOK)
		Me.Panel2.Controls.Add(Me.btnBrowse)
		Me.Panel2.Controls.Add(Me.ScriptInversion)
		Me.Panel2.Dock = System.Windows.Forms.DockStyle.Right
		Me.Panel2.Location = New System.Drawing.Point(232, 0)
		Me.Panel2.Name = "Panel2"
		Me.Panel2.Size = New System.Drawing.Size(230, 30)
		Me.Panel2.TabIndex = 3
		'
		'btnOK
		'
		Me.btnOK.Location = New System.Drawing.Point(150, 0)
		Me.btnOK.Name = "btnOK"
		Me.btnOK.TabIndex = 2
		Me.btnOK.Text = "&OK"
		'
		'btnBrowse
		'
		Me.btnBrowse.Location = New System.Drawing.Point(10, 0)
		Me.btnBrowse.Name = "btnBrowse"
		Me.btnBrowse.Size = New System.Drawing.Size(30, 23)
		Me.btnBrowse.TabIndex = 1
		Me.btnBrowse.Text = "..."
		'
		'Picture
		'
		Me.Picture.BackColor = System.Drawing.Color.White
		Me.Picture.Dock = System.Windows.Forms.DockStyle.Fill
		Me.Picture.Image = CType(resources.GetObject("Picture.Image"), System.Drawing.Image)
		Me.Picture.Location = New System.Drawing.Point(0, 0)
		Me.Picture.Name = "Picture"
		Me.Picture.Size = New System.Drawing.Size(462, 236)
		Me.Picture.TabIndex = 1
		Me.Picture.TabStop = False
		'
		'ScriptInversion
		'
		Me.ScriptInversion.Location = New System.Drawing.Point(50, 0)
		Me.ScriptInversion.Name = "ScriptInversion"
		Me.ScriptInversion.TabIndex = 3
		Me.ScriptInversion.Text = "Script Inversion"
		'
		'Form1
		'
		Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
		Me.ClientSize = New System.Drawing.Size(462, 266)
		Me.Controls.Add(Me.Picture)
		Me.Controls.Add(Me.Panel1)
		Me.Name = "Form1"
		Me.Text = "Image 2 HTML"
		Me.Panel1.ResumeLayout(False)
		Me.Panel2.ResumeLayout(False)
		Me.ResumeLayout(False)

	End Sub

#End Region

	Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
		If Me.OpenFileDialog1.ShowDialog(Me) = DialogResult.OK Then
			Me.FileName.Text = Me.OpenFileDialog1.FileName
			Me.Picture.Image = New Bitmap(Me.OpenFileDialog1.FileName)
		End If
	End Sub

	Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

		Dim HTML As New System.Text.StringBuilder("")
		Dim Height As Integer = Me.Picture.Image.Height
		Dim Width As Integer = Me.Picture.Image.Width

		Dim Image As New Bitmap(Me.Picture.Image)

		HTML.Append("<HTML>")
		HTML.Append("<HEAD>")
		HTML.Append("<TITLE>Image as HTML</TITLE>")
		HTML.Append("<META name=""Generator"" content=""Lewies Image to HTML utility"">")
		HTML.Append("</HEAD>")
		HTML.Append("<BODY marginwidth=0 marginheight=0 topmargin=0 leftmargin=0 rightmargin=0 bottommargin=0>")
		HTML.Append("<TABLE width=" & Width.ToString() & " height=" & Height.ToString() & " cellpadding=0 cellspacing=0 border=0>")
		For Y As Integer = 0 To Height - 1
			HTML.Append("<TR>")
			For X As Integer = 0 To Width - 1
				Dim Color As System.Drawing.Color = Image.GetPixel(X, Y)
				HTML.Append("<TD bgColor=")
				HTML.Append(Hex(Color.R).PadLeft(2))
				HTML.Append(Hex(Color.G).PadLeft(2))
				HTML.Append(Hex(Color.B).PadLeft(2))
				HTML.Append("></TD>")
			Next
			HTML.Append("</TR>")
		Next
		HTML.Append("</TABLE>")

		' Do inversion script
		If Me.ScriptInversion.Checked Then
			HTML.Append("<SCRIPT>")
			HTML.Append(vbCrLf)
			HTML.Append("var cells = new Array();")
			HTML.Append(vbCrLf)
			HTML.Append("window.onload = beginInvert;")
			HTML.Append(vbCrLf)
			HTML.Append("function beginInvert()")
			HTML.Append(vbCrLf)
			HTML.Append("{")
			HTML.Append(vbCrLf)
			HTML.Append("var count = document.all.tags(""TD"").length;")
			HTML.Append(vbCrLf)
			HTML.Append("for(var index=0; index < count; index++)")
			HTML.Append(vbCrLf)
			HTML.Append("{")
			HTML.Append(vbCrLf)
			HTML.Append("cells[index] = index;")
			HTML.Append(vbCrLf)
			HTML.Append("}")
			HTML.Append(vbCrLf)
			HTML.Append("window.setTimeout(""invert()"", 0);")
			HTML.Append(vbCrLf)
			HTML.Append("}")
			HTML.Append(vbCrLf)
			HTML.Append("function invert()")
			HTML.Append(vbCrLf)
			HTML.Append("{")
			HTML.Append(vbCrLf)
			HTML.Append("var index = Math.floor(Math.random() * cells.length);")
			HTML.Append(vbCrLf)
			HTML.Append("var TD = document.all.tags(""TD"")[cells[index]];")
			HTML.Append(vbCrLf)
			HTML.Append("TD.bgColor = inverse(TD.bgColor);")
			HTML.Append(vbCrLf)
			HTML.Append("cells.splice(index, 1);")
			HTML.Append(vbCrLf)
			HTML.Append("if(cells.length != 0)")
			HTML.Append(vbCrLf)
			HTML.Append("{")
			HTML.Append(vbCrLf)
			HTML.Append("window.setTimeout(""invert()"", 0);")
			HTML.Append(vbCrLf)
			HTML.Append("}")
			HTML.Append(vbCrLf)
			HTML.Append("}")
			HTML.Append(vbCrLf)
			HTML.Append("function inverse(color)")
			HTML.Append(vbCrLf)
			HTML.Append("{")
			HTML.Append(vbCrLf)
			HTML.Append("if(color.length == 7)")
			HTML.Append(vbCrLf)
			HTML.Append("{")
			HTML.Append(vbCrLf)
			HTML.Append("color = color.substring(1, 7);")
			HTML.Append(vbCrLf)
			HTML.Append("}")
			HTML.Append(vbCrLf)
			HTML.Append("color = ""00000"" + (0xffffff - parseInt(color, 16)).toString(16);")
			HTML.Append(vbCrLf)
			HTML.Append("return color.substr(color.length - 6, 6);")
			HTML.Append(vbCrLf)
			HTML.Append("}")
			HTML.Append(vbCrLf)
			HTML.Append("</SCRIPT>")
		End If

		HTML.Append("</BODY>")

		' save to file
		Dim FileName = Application.StartupPath & "\preview.htm"
		Dim buffer As Byte()
		buffer = System.Text.UnicodeEncoding.Unicode.GetBytes(HTML.ToString)
		Dim File As System.IO.FileStream = New System.IO.FileStream(FileName, IO.FileMode.Create)
		File.Write(buffer, 0, buffer.Length)
		File.Close()
		File = Nothing

		' create container
		HTML = New System.Text.StringBuilder("")
		HTML.Append("<HTML>")
		HTML.Append("<HEAD>")
		HTML.Append("<TITLE>Image 2 HTML</TITLE>")
		HTML.Append("</HEAD>")
		HTML.Append("<BODY>")
		HTML.Append("<CENTER>")
		HTML.Append("<H1>Image 2 HTML Example</H1>")
		HTML.Append("<IFRAME frameborder=""no"" src=""preview.htm""")
		HTML.Append(" width=""" & Me.Picture.Image.Width & """")
		HTML.Append(" height=""" & Me.Picture.Image.Height & """")
		HTML.Append("></IFRAME>")
		If Me.ScriptInversion.Checked Then
			HTML.Append("<P>Watch as the image inverts itself</P>")
		End If
		HTML.Append("<P>Utility created by <A href=""http://www.lewismoten.com"">Lewis Moten</A></P>")
		HTML.Append("</CENTER>")
		HTML.Append("</BODY>")

		' save to file
		FileName = Application.StartupPath & "\previewContainer.htm"
		buffer = System.Text.UnicodeEncoding.Unicode.GetBytes(HTML.ToString)
		File = New System.IO.FileStream(FileName, IO.FileMode.Create)
		File.Write(buffer, 0, buffer.Length)
		File.Close()
		File = Nothing

		' open in browser
		Shell("explorer.exe " & FileName)

	End Sub
End Class
