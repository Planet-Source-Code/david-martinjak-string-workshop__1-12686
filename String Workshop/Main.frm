VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "String Workshop"
   ClientHeight    =   2190
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   4680
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4680
   Begin VB.TextBox txtBody 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Main.frx":08CA
      ToolTipText     =   "Type in here"
      Top             =   0
      Width           =   4695
   End
   Begin VB.Menu mnuStrings 
      Caption         =   "Strings"
      Begin VB.Menu mnuConcatenate 
         Caption         =   "Concatenate"
      End
      Begin VB.Menu mnuExtract 
         Caption         =   "Extract"
         Begin VB.Menu mnuExtractLeft 
            Caption         =   "Left"
         End
         Begin VB.Menu mnuExtractMiddle 
            Caption         =   "Middle"
         End
         Begin VB.Menu mnuExtractRight 
            Caption         =   "Right"
         End
      End
      Begin VB.Menu mnuReplace 
         Caption         =   "Replace"
      End
      Begin VB.Menu mnuLine0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnuLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu mnuLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAboutMsg 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'String Workshop
'Author:  David Martinjak  ·aka·  pancho
'Date:    November 10, 2000
'Purpose: To help novices learn about string routines


'for loading web browser
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
"ShellExecuteA" (ByVal hwnd As Long, _
ByVal lpOperation As String, ByVal lpFile As String, _
ByVal lpParameters As String, ByVal lpDirectory As String, _
ByVal nShowCmd As Long) As Long



 
Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Private Sub mnuAboutMsg_Click()
MsgBox "String Workshop" & vbCrLf & "© David Martinjak" & vbCrLf & "513 Designs, Inc. ®", vbInformation, "String Workshop"
End Sub

Private Sub mnuConcatenate_Click()
Dim newString As String, oldLength As Integer

'first we ask the user what new string he or she wants to
'concatenate.  by the way:  concatenation is a big tech term
'for adding.  you're just adding one string on to another,
'don't get intimidated by terminology ;]

newString = InputBox("What would you like to add to the text:", "String Workshop")

'check to see the user didn't just leave it blank
If newString = "" Then Exit Sub

'this is just to see how many characters are in the textbox
'before adding the value of newString to it
oldLength = Len(txtBody)

'now we add on the new string
txtBody = txtBody & newString

'                -^-
'notice the use of the ampersand.  you can also use a plus
'sign (+), but i prefer to reserve those for math computation

'now we'll select the new text so you can see the addition
'you've made
txtBody.SelStart = oldLength
txtBody.SelLength = Len(newString)
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuExtractLeft_Click()
Dim extract As String

'this will use the Left function to extract the first
'n amount of characters specified by the user.
extract = InputBox("how many characters would you like to extract from the left:", "String Workshop")

'tell the user what the extracted text is
'the Left function first asks what string you'd like to
'examine, then how many characters you'd like to extract
'starting with the very first.

MsgBox "The first " & extract & " characters of the text are:" _
       & vbCrLf & Left(txtBody, Int(Val(extract))), vbInformation, "String Workshop"
       
'select the extracted text
txtBody.SelStart = 0
txtBody.SelLength = Int(Val(extract))
End Sub

Private Sub mnuExtractMiddle_Click()
Dim start As String, length1 As String, extract As String

'in this sub, we're going to extract characters from the
'middle of the string.  we'll be using the Mid function
'to complete the task.  now the Mid function was a bit
'difficult to understand and master for me personally so
'i'm going to explain it as best i can.

'the declaration for Mid looks like such:
'Function Mid(String, Start As Long, [Length])

'let's break this down and look at it
'first parameter:
'   it wants to know what string to look at, ok simple enough

'second parameter:
'   it wants to know which character to start at.
'   unlike arrays, strings start with a value of 1.
'   so if you want to extract all characters from the 4th
'   to the last, your code would look like this:
'   extract = Mid(txtBody, 4)

'third parameter:
'   notice that Length is in brackets [ ].
'   this means it's optional.
'   you don't have to fill in anything there, but if you
'   choose not to, it will extract from the start to the end
'   of the string.
'   say you want to extract characters 4, 5, and 6; your
'   coding would look like this:
'   extract = Mid(txtBody, 4, 3)

start = InputBox("which character would you like to start at:", "String Workshop")
If start = "" Then Exit Sub

length1 = InputBox("how many characters would you like to extract (optional):", "String Workshop")
If length1 = "" Then length1 = "\0" 'set it to null

If length1 = "\0" Then
  extract = Mid(txtBody, CLng(start))
  MsgBox "extracted text:" & vbCrLf & extract, vbInformation, "String Workshop"
  
  'select extracted text
  txtBody.SelStart = CInt(start)
  txtBody.SelLength = Len(txtBody) - CInt(start)
Else
  extract = Mid(txtBody, CLng(start), CLng(length1))
  MsgBox "extracted text:" & vbCrLf & extract, vbInformation, "String Workshop"
  
  'select extracted text
  txtBody.SelStart = CInt(start)
  txtBody.SelLength = CInt(length1)
End If
    
End Sub


Private Sub mnuExtractRight_Click()
Dim extract As String

'this operation is very similar to the Left function, except
'that it starts from the last character and works its way to
'the beginning of the string.

extract = InputBox("how many characters would you like to extract:", "String Workshop")
If extract = "" Then Exit Sub

'extract text and notify user
MsgBox extract & " characters from the last character is:" _
    & vbCrLf & Right(txtBody, CLng(extract)), vbInformation, "String Workshop"

'select text
txtBody.SelStart = Len(txtBody) - CInt(extract)
txtBody.SelLength = CInt(extract)
End Sub

Private Sub mnuHelp_Click()
Call ShellExecute(Me.hwnd, "Open", "http://www.stas.net/5/ruin", vbNullString, vbNullString, 3)
End Sub

Private Sub mnuOpen_Click()
Dim Path As String

'ask user for the text file
'using a do-while loop continues to ask the user until they
'specify a .txt file or cancel

Do
 Path = InputBox("location of text (*.txt) file:", "String Workshop [Open]", App.Path & "\string.txt")
 If Path = "" Then Exit Sub
Loop While Right(Path, 4) <> ".txt"

'open the text file to txtBody
Open Path For Input As 1&
    Let txtBody = Input$(LOF(1&), 1&)
Close 1&
End Sub


Private Sub mnuReplace_Click()
Dim oldString As String, newString As String

'replacing characters or even strings within a string is not
'difficult at all.  vb comes with a Replace function built
'right into it so you don't even have to code yourself
'anything here.  remember however that when you specifiy
'what you want to replace, it is case sensitive.

oldString = InputBox("what character or string do you want to replace:", "String Workshop")
newString = InputBox("what do you want to replace " & oldString & " with: ", "String Workshop")

'with Replace, you just tell it what string to look at,
'in this case the text in txtBody.  then you tell it what
'string you want out, and what string you want in it's place

txtBody = Replace(txtBody, oldString, newString)
End Sub


Private Sub mnuSave_Click()
Dim Path As String

'ask user for the text file
'using a do-while loop continues to ask the user until they
'specify a .txt file or cancel

Do
 Path = InputBox("location of text (*.txt) file:", "String Workshop [Save]", App.Path & "\string.txt")
 If Path = "" Then Exit Sub
Loop While Right(Path, 4) <> ".txt"

'save the text file
Open Path For Output As #1
    Print #1, txtBody
Close #1
End Sub

Private Sub mnuStrings_Click()
'if there's nothing to save, don't give them that option
If Len(txtBody) = 0 Then
  mnuSave.Enabled = False
Else
  mnuSave.Enabled = True
End If
End Sub


