= StringBuffer Class 

	This class provides fast string of charactors.

= Author

	kazuya kawaguchi (kawakazu80@gmail.com)


= Installation

	1. Download the source repository from 'git@github.com:kazupon/StringBuffer4VB6.git'.

	2. Add 'StringBuffer.cls' file on Visual Basic 6 project.


= Usage

	Dim objStringBuffer As StringBuffer
	Set objStringBuffer = New StringBuffer ' create object
	
	Call objStringBuffer.Append("hello world.") ' append string
	Call objStringBuffer.Appends("a", "b", "c") ' apeend strings
	Call objStringBuffer.AppendLine("hello world.") ' append string + vbCrLf
	Call objStringBuffer.AppendsLine(""a", "b", "c") ' append stirngs + vbCrLf
	
	Dim strMessage As String
	strMessage = objStringBuffer.ToString() ' buffered strings to string
	Debug.Print strMessage
	
	Call objStringBuffer.Clear	' clear buffered strings

