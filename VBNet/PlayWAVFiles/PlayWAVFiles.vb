'**************************************
' Name: Play WAV Files
' Description:Allows you to effortlessly play wave files in a seperate thread while your application continues to operate. Now supports playing embedded resource files and Streams.
' By: Lewis Moten
'
'
' Inputs:None
'
' Returns:None
'
'Assumes:None
'
'Side Effects:None
'This code is copyrighted and has limited warranties.
'**************************************

Option Strict On
Option Explicit On 
Imports System
Imports System.Threading
Public Class clsAudio
	'     ' Author: Lewis Moten
	'     ' URL: http:'www.lewismoten.com
	'     ' Date: 02-15-2003
	'     ' Examples
	'Dim Audio As New clsAudio("shuffle.wav"
	'     )
	'Audio.Play("shuffle.wav", clsAudio.enSo
	'     und.Memory)
	'     'Audio.Play(Stream)
	Public Enum enSound
		Application = &H80L
		[Alias] = &H10000L
		AliasID = &H110000L
		Async = &H1L
		FileName = &H20000L
		[Loop] = &H8L
		Memory = &H4L
		NoDefault = &H2L
		NoStop = &H10L
		NoWait = &H2000L
		Purge = &H40L
		Resource = &H40004L
		Sync = &H0L
	End Enum
	'     ' If flag not set, default is used
	Private DefaultFlags As enSound = enSound.Application
	Public Sub New()
	End Sub
	#Region " Play Stream "
	Public Sub Play(ByVal Stream As System.IO.Stream)
		PlayThread(Stream, DefaultFlags Xor enSound.Memory)
	End Sub
	Public Sub Play(ByVal Stream As System.IO.Stream, ByVal Flags As enSound)
		PlayThread(Stream, DefaultFlags Xor enSound.Memory)
	End Sub
	Public Sub New(ByVal Stream As System.IO.Stream)
		PlayThread(Stream, DefaultFlags Or enSound.Memory)
	End Sub
	Public Sub New(ByVal Stream As System.IO.Stream, ByVal Flags As enSound)
		PlayThread(Stream, DefaultFlags Or enSound.Memory)
	End Sub
	#End Region
	#Region " Play File "
	Public Sub Play(ByVal FileName As String)
		PlayThread(FileName, DefaultFlags)
	End Sub
	Public Sub Play(ByVal FileName As String, ByVal Flags As enSound)
		PlayThread(FileName, Flags)
	End Sub
	Public Sub New(ByVal FileName As String)
		PlayThread(FileName, DefaultFlags)
	End Sub
	Public Sub New(ByVal FileName As String, ByVal Flags As enSound)
		PlayThread(FileName, Flags)
	End Sub
	#End Region
	#Region "PlayThread"
	Private Sub PlayThread(ByVal FileName As String, ByVal Flags As enSound)
		Dim oPlayer As New clsPlayer()
		oPlayer.FileName = FileName
		oPlayer.Flags = Flags
		Dim oThread As New Thread(AddressOf oPlayer.Play)
		oThread.Start()
	End Sub
	Private Sub PlayThread(ByVal Stream As System.IO.Stream, ByVal Flags As enSound)
		Dim oPlayer As New clsPlayer()
		oPlayer.Stream = Stream
		oPlayer.Flags = Flags
		Dim oThread As New Thread(AddressOf oPlayer.Play)
		oThread.Start()
	End Sub
	'     ' Internal class used for threading.
	Private Class clsPlayer
		Public FileName As String
		Public Flags As enSound
		Public Stream As System.IO.Stream
		Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySound" (ByVal pszSound() As Byte, ByVal hMod As Int16, ByVal fdwSound As Long) As Boolean
		Private Sub LoadResource()
			If FileName = "" Then Exit Sub
			Dim oAssembly As Reflection.Assembly = MyClass.GetType.Assembly
			Dim name As String = oAssembly.GetName.Name
			name &= "." & FileName
			Stream = oAssembly.GetManifestResourceStream(name)
		End Sub
		Public Sub Play()
			'http:'msdn.microsoft.com/library/defau
			'     lt.asp?url=/library/en-us/multimed/mmfun
			'     c_9uxw.asp
			Dim Success As Boolean ' Return Code
			Dim Path As String = System.AppDomain.CurrentDomain.BaseDirectory
			Dim FullPath As String
			Dim buffer() As Byte
			If CBool(Flags And enSound.Memory) Then
				' User says they wish to play file in me
				'     mory
				If Stream Is Nothing Then
					' Stream isn't loaded - lets load that f
					'     ile
					' for the user from our embedded resourc
					'     es
					' Check to make sure they specified the 
					'     audio to be played
					If FileName = "" Then
						Throw New Exception("Filename not provided")
						Return
					End If
					'     ' Load that file
					LoadResource()
					If Stream Is Nothing Then
						'     ' Woah - that file didn't exist as a
						'     ' resource. Let's tell them about it.
						Throw New Exception("File """ & FileName & """ does not exist as an embedded resource file.")
						Return
					End If
				End If
			Else
				' User says they wish to play audio, but
				'     that the
				' audio isn't in memory. It is probably 
				'     a file
				'     ' on the file system.
				' Check to make sure they specified the 
				'     audio to be played
				If FileName = "" Then
					Throw New Exception("Filename not provided")
					Return
				End If
				' Hmm... lets see if it is in the curren
				'     t working
				' directory, or in an absolute path on t
				'     he file
				'     ' system.
				If Not InStr(FileName, ":\") = 2 Then
					FullPath = Path & FileName
				Else
					FullPath = FileName
				End If
				' Ok - lets make sure that file actually
				'     exists.
				If Not System.IO.File.Exists(FullPath) Then
					' Nope - let's assume it is an embedded 
					'     resource
					LoadResource()
					'     ' were we able to load it?
					If Stream Is Nothing Then
						'     ' Nope - not an embedded resource.
						Throw New Exception("File """ & FileName & """ does not exist on the file system or as an embedded resource file.")
						Return
					End If
					' Make sure we tell the computer that ou
					'     r stream is
					'     ' an audio file in memory
					Flags = Flags Or enSound.Memory
				End If
			End If

			If Not CBool(Flags And enSound.Memory) Then
				' We are playing a file on the file syst
				'     em ...
				'     ' buffer the file name
				buffer = System.Text.ASCIIEncoding.ASCII.GetBytes(FullPath)
			Else
				'     ' We are playing a file from memory
				'     ' buffer the stream
				ReDim buffer(CType(Stream.Length, Integer))
				Stream.Read(buffer, 0, buffer.Length)
			End If
			'     ' Play that sound
			Success = PlaySound(buffer, 0, Flags)
			'     ' Did we play it?
			If Not Success Then
				'     ' Nope - let's tell the user.
				If FileName = "" Then
					Throw New Exception("Audio stream could not be played")
				Else
					Throw New Exception("Audio """ & FileName & """ could not be played")
					End If
				Return
			End If
		End Sub
	End Class
	#End Region
End Class