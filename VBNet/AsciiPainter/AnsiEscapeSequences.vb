Public Enum AnsiEscapeSequence

	InsertCharacter = 64	' @

	' A-Z (65-90)
	MoveCursorUp = 65
	MoveCursorDown = 66
	MoveCursorRight = 67
	MoveCursorLeft = 68
	LineFeed = 69
	UndefinedF = 70
	UndefinedG = 71
	MoveCursorTo = 72
	UndefinedI = 73
	ClearScreen = 74
	ClearToEndOfLine = 75
	InsertLine = 76
	PlayMusic = 77	'  (Originally used for DeleteLine, BBS use it for music)
	PlayMusicFix = 78	' N Fixes M
	UndefinedO = 79
	DeleteCharacter = 80
	UndefinedQ = 81
	UndefinedR = 82
	ScrollUp = 83
	ScrollDown = 84
	Clear = 85
	UndefinedV = 86
	UndefinedW = 87
	UndefinedX = 88
	DeleteLine = 89	' Y Fixes M
	BackTab = 90

	' a-z (97-122)
	Undefined_a = 97
	Undefined_b = 98
	ReportIsConnected = 99
	Undefined_d = 100
	Undefined_e = 101
	MoveCursorTo2 = 102
	Tabs = 103	' Unknown
	SetOptions = 104
	Print = 105
	Undefined_j = 106
	Undefined_k = 107
	SetOptions2 = 108
	SetVideoOptions = 109
	ReportDisplayInformation = 110
	Undefined_o = 111
	ReassignKeyboard = 112
	SetKeyboardOptions = 113
	SetScroolRegion = 114
	SaveCursorPosition = 115
	Undefined_t = 116
	RestoreCursorPosition = 117
	Undefined_v = 118
	Undefined_w = 119
	Undefined_x = 120
	Test = 121
	Reset = 122
End Enum
