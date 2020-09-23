<div align="center">

## Frequently Asked VB Questions


</div>

### Description

If you've ever been to the Visual Basic Discussion Forum then you realize why I'm posting this. It's simply a list of commonly asked questions. Please feel free to add on anything that either you've learned here at PSC or anything that you find yourself answering on a regular basis. (Revised Jul 06, 2001)

Revised Jul 19, 2001

Revised Jul 24, 2001
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Sean Street](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sean-street.md)
**Level**          |Beginner
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, VBA MS Access, VBA MS Excel
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/sean-street-frequently-asked-vb-questions__1-24689/archive/master.zip)





### Source Code

<table border=2>
<tr>
<td><b><center>Functionality</center></b></td>
<td><b><center>Relative Code</center></b></td>
<td><b><center>Related Links</center></b></td>
</tr>
<tr>
<td>Writing/Appending text to a text file</td>
<td><pre>
Open "C:\MyTextFile.txt" For Output As #1<br>Open "C:\MyTextFile.txt" For Append As #1
</pre>
</td>
<td><a href="http://www.planet-source-code.com/xq/ASP/txtCodeId.22246/lngWId.1/qx/vb/scripts/ShowCode.htm">Input/Output Text file</a></td>
</tr>
<tr>
<td>Reading text from a text file</td>
<td><pre>
Open "C:\MyTextFile.txt" For Input As #1
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/xq/ASP/txtCodeId.22246/lngWId.1/qx/vb/scripts/ShowCode.htm">Input/Output Text file</a></td>
</tr>
<tr>
<td>Setting a string to the application directory</td>
<td><pre>
strFileName = App.Path & (Trim(Chr(32 - (60 * (Asc(Right(App.Path, 1)) <> 92)))))
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/vb/Discussion/AskAProShowPost.asp?lngTopicId=10826&Forum=Visualbasic&TopicCategory=programming&Flag=2&lngWId=1">Relative paths</a></td>
</tr>
<tr>
<td>Reading data from an INI file</td>
<td><pre>
Private Declare Function GetPrivateProfileString Lib "kernel32"_<br> Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any,_<br> ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long,_<br> ByVal lpFileName As String) As Long
<br><br>
Public Function GetINIData(ByVal strParent As String, strKey As String) As String<br>
  Dim strBuffer As String<br>
  Dim strFilename As String<br><br>
  strBuffer = Space(145)<br>
  strFileName = App.Path & (Trim(Chr(32 - (60 * (Asc(Right(App.Path, 1)) <> 92))))) & "MyINI.INI"<br><br>
  GetPrivateProfileString strParent, strKey, "", strBuffer, Len(strBuffer) - 1, strFilename<br>
  GetINIData = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)<br>
End Function<br>
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/xq/ASP/txtCodeId.23487/lngWId.1/qx/vb/scripts/ShowCode.htm">INI file template routines</a></td>
</tr>
<tr>
<td>Writing data to an INI file</td>
<td><pre>
Private Declare Function WritePrivateProfileString Lib "kernel32"_<br>
Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String,_<BR>
ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
<br><br>
Public Sub WriteINIData(ByVal strParent As String, strKey As String, strValue As String)<br>
  Dim strFilename As String<br><br>
  strFileName = App.Path & (Trim(Chr(32 - (60 * (Asc(Right(App.Path, 1)) <> 92))))) & "MyINI.INI"<br><br>
  WritePrivateProfileString strParent, strKey, strValue, strFilename<br>
End Sub<br>
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/xq/ASP/txtCodeId.23487/lngWId.1/qx/vb/scripts/ShowCode.htm">INI file template routines</a></td>
</tr>
<tr>
<td>Dynamically adding controls</td>
<td><pre>
Rem This code is for Visual Basic 6 only but the second link shows how to do it with VB4/5<br>
Private Sub Form_Load()<br>
Form1.Controls.Add "VB.CommandButton", "cmdMyButton"<br>
With Form1!cmdMyButton<br>
.Visible = True<br>
.Width = 2000<br>
.Caption = "Dynamic Button"<br>
End With<br>
End Sub<br>
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/vb/discussion/AskAProShowPost.asp?lngWId=1&Flag=2&TopicCategory=programming&lngTopicId=4870&Forum=Visualbasic">Dynamically create a control(VB6)</a><br><br>
<a href="http://www.planet-source-code.com/vb/discussion/AskAProShowPost.asp?lngWId=1&Flag=2&TopicCategory=programming&lngTopicId=5147&Forum=Visualbasic">Creating controls dynamically (VB6,5 and 4)</a></td>
</tr>
<tr>
<td>Adding items to a combo/list box and<br>setting it to the first items if an item exist</td>
<td><pre>
cmbMyComboBox.AddItem "Item1"<br>
cmbMyComboBox.ListIndex = (cmbMyComboBox.ListCount=0)
</pre></font>
</td>
<td>None</td>
</tr>
<tr>
<td>Having problems with the license of your Winsock control?</td>
<td><pre>
Just go to the link
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/xq/ASP/txtCodeId.4860/lngWId.1/qx/vb/scripts/ShowCode.htm">Register/License Winsock Control</td>
</tr>
<tr>
<td>Allows only numeric characters in a textbox</td>
<td><pre>
Private Sub txtNumbersOnly_KeyPress(KeyAscii As Integer)<br>
  KeyAscii = KeyAscii * Abs(((KeyAscii > 47 And KeyAscii < 58) Or KeyAscii = vbKeyBack))<br>
End Sub<br>
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/xq/ASP/txtCodeId.11545/lngWId.1/qx/vb/scripts/ShowCode.htm">Masking Control</td>
</tr>
<tr>
<td>Prints a picture control contents to the printer</td>
<td><pre>
Printer.PaintPicture picMyPictureControl.Picture, 1, 1
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/vb/discussion/AskAProShowPost.asp?lngWId=1&Flag=2&TopicCategory=standards&lngTopicId=10496&Forum=Visualbasic">Printing picture control contents</td>
</tr>
<tr>
<td>Copy picture/text to the Clipboard</td>
<td><pre>
Clipboard.Clear
Clipboard.SetData picMyPictureControl.Picture 'Used for pictures<br>
Clipboard.SetText txtMyTextBox.Text 'Used for text<br>
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/vb/discussion/AskAProShowPost.asp?lngWId=1&Flag=2&TopicCategory=standards&lngTopicId=9503&Forum=Visualbasic">Copying contents to the Clipboard</td>
</tr>
<tr>
<td>Paste picture/text from the Clipboard</td>
<td><pre>
picMyPictureControl.Picture = Clipboard.GetData 'Used for pictures<br>
txtMyTextBox.Text = Clipboard.GetText 'Used for text<br>
</pre></font>
</td>
<td><a href="http://www.planet-source-code.com/vb/discussion/AskAProShowPost.asp?lngWId=1&Flag=2&TopicCategory=standards&lngTopicId=9503&Forum=Visualbasic">Pasting contents from the Clipboard</td>
</tr>
<tr>
<td>Evaluate resposes from MsgBox</td>
<td><pre>
Rem Use this to check before you save; used with yes/no or ok/cancel options<br>
If MsgBox("Are you sure you want to save thses changes?", vbQuestion + vbYesNo, "Save?") = vbNo Then Exit Sub<br>
<br>
<br>
Rem You can use this to check before you exit; used with yes/no/cancel or abort/retry/ignore<br>
Select Case MsgBox("Would you like to save before you exit?", vbQuestion + vbYesNoCancel, "Exiting")<br>
Case vbYes<br>
Rem Save it then quit<br>
Case vbNo<br>
Rem Quit<br>
Case vbCancel<br>
Exit Sub<br>
End Select<br><br>
</pre></font>
</td>
<td>None</td>
</tr>
<tr>
<td>Read data from an Excel spreadsheet</td>
<td><pre>
Dim xlsApplication As Object<br>
Dim lngRowCount As Long<br>
Dim intColCount As Integer<br>
Dim blnBlankRow As Boolean<br>
Dim strValue As String<br>
<br>
Set xlsApplication = CreateObject("Excel.Application")<br>
<br>
xlsApplication.Workbooks.Open "C:\Test.XLS"<br>
<br>
For lngRowCount = 1 To 65536<br>
	blnBlankRow = True<br>
	For intColCount = 1 To 255<br>
		strValue = xlsApplication.Cells(lngRowCount, intColCount).Value<br>
		Rem Set this value into your table/field<br>
		If Len(strValue) > 0 Then blnBlankRow = False<br>
	Next intColCount<br>
	If blnBlankRow Then Exit For<br>
Next lngRowCount<br>
<br>
xlsApplication.Workbooks(1).Close savechanges:=False<br>
xlsApplication.Quit<br><br>
</pre></font>
</td>
<td>None</td>
</tr>
<tr>
<td>Read data from Outlook Inbox/SentMail folders</td>
<td><pre>
Dim outApplication As Object<br>
Dim outInBox As Object<br>
Dim outOutBox As Object<br>
<br>
Set outApplication = CreateObject("Outlook.Application")<br>
<br>
Set outInBox = outApplication.GetNamespace("MAPI").GetDefaultFolder(6)<br>
Set outOutBox = outApplication.GetNamespace("MAPI").GetDefaultFolder(5)<br>
<br>
Rem First InBox email<br>
MsgBox outInBox.Items.Item(1).Recipients(1).Name, vbOKOnly, "Inbox Recipient"<br>
MsgBox outInBox.Items.Item(1).Subject, vbOKOnly, "Inbox Subject"<br>
MsgBox outInBox.Items.Item(1).Body, vbOKOnly, "Inbox Body"<br>
<br>
Rem First SentMail email<br>
MsgBox outOutBox.Items.Item(1).Recipients(1).Name, vbOKOnly, "SentMail Recipient"<br>
MsgBox outOutBox.Items.Item(1).Subject, vbOKOnly, "SentMail Subject"<br>
MsgBox outOutBox.Items.Item(1).Body, vbOKOnly, "SentMail Body"<br><br>
</pre></font>
</td>
<td>None</td>
</tr>
<tr>
<td>Sending email using the MS Outlook object</td>
<td><pre>
Private Sub MrPostman(strSendTo As String, strSubject As String, strMessage As String)<br>
  Dim outEmail As Outlook.Application<br>
  Dim outNewMail As Outlook.MailItem<br>
  Dim strTemp() As String<br>
<br>
  Set outEmail = New Outlook.Application<br>
  Set outNewMail = outEmail.CreateItem(olMailItem)<br>
<br>
  With outNewMail<br>
<br>
    strTemp = Split(strSendTo, ";")<br>
<br>
    For intCounter = 0 To UBound(strTemp)<br>
      .Recipients.Add Trim(strTemp(intCounter))<br>
    Next intCounter<br>
<br>
    .Subject = strSubject<br>
    .Body = strMessage<br>
    .Send<br>
  End With<br>
<br>
  Set outEmail = Nothing<br>
  Set outNewMail = Nothing<br>
<br>
End Sub<br>
</pre></font>
</td>
<td>None</td>
</tr>
<tr>
<td>Calling procedures dynamically</td>
<td><pre>
Rem Use this code when you don't know the name of the procedure or when you want the user to select the procedure to execute<br>
Private Sub Form_Load()<br>
CallByName Form1, "Test", VbMethod<br>
End Sub<br>
Public Function Test()<br>
  MsgBox "It Works"<br>
End Function<br>
</pre></font>
</td>
<td>None</td>
</tr>
<tr>
<td>Copy/Move files from one location to another</td>
<td><pre>
FileCopy "C:\SourceFile.txt", "C:\DestinationFile.txt"<br>
Rem To move the file (delete the original)<br>
Kill "C:\SourceFile.txt"<br>
</pre></font>
</td>
<td>None</td>
</tr>
<tr>
<td>Retained is an invalid key error</td>
<td><pre>
You will get this error when you attempt to open a project designed in VB6+ with VB5-.
The solution is to open the project file (*.vbp) with a text editor like notepad
and delete the line that begins with RETAINED=. This will solve the error.
</pre></font>
</td>
<td>None</td>
</tr>
<tr>
<td>What does referencing a control mean?<br>What is the difference between early and late binding?</td>
<td><pre>
When you create a reference to a control, you are indicating that there is a file that exists
that you would like to use. Early-binding indicates this reference at design-time of the application
rather than an runtime (late binding). Early binding is much faster than late binding. Late binding
is used when an application must determine at runtime. Although this process is slower than
late binding, it may be faster after consideration. For example, let's say that you are importing
data from one source to another. You are uncertain at design time wheter the user will want to
import from Excel to Access, Outlook to Excel, Outlook to Access, Excel to Outlook, Access to Outlook,
or Access to Excel. Instead of referencing all three objects at design time(early binding),it may
be more practical to refernce them once the user has mad a decision (late binding).
</pre></font>
</td>
<td>None</td>
</tr>
</table>

