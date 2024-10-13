'**************************************
' Name: Thumbnail
' Description:allows you to dynamically create thumbnails of your images on the fly. Image is resized to its aspect ration to fit within a defined region. You can specify quality of image to reduce bandwidth, and background color to fill in whitespace around image.
' By: Lewis Moten
'Assumes:the user should know how to retrieve binary data of a file from the file system or from a database.
'This code is copyrighted and has limited warranties.
'**************************************

Option Explicit On 
Option Strict On
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO

Public Class thumbnail
    Inherits System.Web.UI.Page
    Dim height As Integer = 96
    Dim width As Integer = 128
    Dim quality As Integer = 10
    Dim background As Color = Color.White
#Region " Web Form Designer Generated Code "
    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
    End Sub
    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object
    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
    	'CODEGEN: This method call is required by the Web Form Designer
    	'Do not modify it using the code editor.
    	InitializeComponent()
    End Sub
#End Region
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
    	' width defines how wide we want our thumbnail to be.
    	Try : width = Integer.Parse(Request.QueryString("width"))
    	Catch ex As Exception : End Try
    	Try : height = Integer.Parse(Request.QueryString("height"))
    	Catch ex As Exception : End Try
    	' quality 1-100 determines how good the quality of the thumbnail
    	' will appear. smaller quality results in less bandwidth, but
    	' with bad quality.
    	Try : quality = Integer.Parse(Request.QueryString("quality"))
    	Catch ex As Exception : End Try
    	' since images are resized "to fit within" the thumbnails dimensions,
    	' there is whitespace left over. You can change the color of this
    	' white space by providing an HTML hex col
    	Dim bgColor As String = Request.QueryString("bgColor")
    	If Not bgColor Is Nothing Then
    		bgColor = bgColor.Replace("#", "")
    		If bgColor.Length = 6 Then
    			Dim red As Integer = CType("&H" & bgColor.Substring(0, 2), Integer)
    			Dim green As Integer = CType("&H" & bgColor.Substring(2, 2), Integer)
    			Dim blue As Integer = CType("&H" & bgColor.Substring(4, 2), Integer)
    			background = background.FromArgb(red, green, blue)
    		End If
    	End If
    End Sub
    Private Function BinaryData() As Byte()
    	' this is where we load our data from the database
    	' as a byte array.
    	' you will need to change this to access your database.
    	Dim item As Integer
    	item = Integer.Parse(Request.QueryString("item"))
    	Dim Gallery As New ImageGallery
    	Dim Table As DataTable = Gallery.Image(item)
    	Dim Row As DataRow = Table.Rows(0)
    	Return CType(Row.Item("binarydata"), Byte())
    End Function
    Private Sub resize(ByRef imageWidth As Integer, ByRef imageHeight As Integer, ByVal maxWidth As Integer, ByVal maxHeight As Integer)
    	' this method will resize your image by its aspect ratio to fit
    	' within the specified boundaries.
    	If imageWidth > maxWidth Then
    		imageHeight = CType(imageHeight * (maxWidth / imageWidth), Integer)
    		imageWidth = maxWidth
    	End If
    	If imageHeight > maxHeight Then
    		imageWidth = CType(imageWidth * (maxHeight / imageHeight), Integer)
    		imageHeight = maxHeight
    	End If
    End Sub
    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
    	' this method actually resizes and writes the image to the users
    	' browser.
    	' load the image from binary data
    	Dim image As Drawing.Image
    	image = image.FromStream(New MemoryStream(BinaryData()))
    	' determine internal size of thumbnail by resizing image
    	' by aspect ration to fit within specified region
    	Dim thumbWidth As Integer = image.Width
    	Dim thumbHeight As Integer = image.Height
    	resize(thumbWidth, thumbHeight, width, height)
    	' Setup graphics object to work with thumbnail image
    	Dim thumbnail As System.Drawing.Image = New Bitmap(width, height)
    	Dim g As Graphics = g.FromImage(thumbnail)
    	' setup background color of thumbnail
    	g.FillRectangle(New SolidBrush(background), 0, 0, width, height)
    	' find offsets based on space available
    	Dim offsetX As Integer = CType((width - thumbWidth) / 2, Integer)
    	Dim offsetY As Integer = CType((height - thumbHeight) / 2, Integer)
    	' draw the resized image in the center of the canvis
    	g.DrawImage(image, offsetX, offsetY, thumbWidth, thumbHeight)
    	' setup quality of image to save
    	Dim Info() As ImageCodecInfo = ImageCodecInfo.GetImageEncoders
    	Dim Params As New EncoderParameters(1)
    	Params.Param(0) = New EncoderParameter(Encoder.Quality, quality)
    	' send the thumbnail to the users browser
    	Response.ContentType = Info(1).MimeType
    	thumbnail.Save(Response.OutputStream, Info(1), Params)
    End Sub
End Class