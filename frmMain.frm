VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Advanced Command Processor/Parser - Jamie S"
   ClientHeight    =   8745
   ClientLeft      =   1905
   ClientTop       =   1935
   ClientWidth     =   6135
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8745
   ScaleWidth      =   6135
   Begin VB.TextBox Text4 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1335
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "frmMain.frx":1CFA
      Top             =   7320
      Width           =   5655
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   3015
      Left            =   3960
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   13
      Text            =   "frmMain.frx":1E9D
      Top             =   4080
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   3960
      Picture         =   "frmMain.frx":1FDB
      ScaleHeight     =   720
      ScaleWidth      =   720
      TabIndex        =   11
      Top             =   3240
      Width           =   720
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   9
      Top             =   4200
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Execute"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   3240
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Text            =   "frmMain.frx":2895
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\\ = \ (backslash)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3960
      TabIndex        =   15
      Top             =   1080
      Width           =   1485
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   1455
      Left            =   120
      Top             =   7200
      Width           =   5895
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Advanced Command Processor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   675
      Left            =   4800
      TabIndex        =   12
      Top             =   3240
      Width           =   1080
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   3975
      Left            =   3840
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "How the processor interprets it..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   3960
      Width           =   2400
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\t = tab (vbTab)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3960
      TabIndex        =   8
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Tip: Enclose an argument in quotes ("""") if you do not want leading and trailing white space trimmed i.e.   ""  argument  """
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   1215
      Left            =   3960
      TabIndex        =   7
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\"" = "" (double-quote)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3960
      TabIndex        =   6
      Top             =   840
      Width           =   1740
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\n = newline (vbCrLf)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3960
      TabIndex        =   5
      Top             =   1560
      Width           =   1755
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\; = ; (colon)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3960
      TabIndex        =   4
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "\, = , (comma)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   3960
      TabIndex        =   3
      Top             =   360
      Width           =   1245
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Valid Escape Characters"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   210
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   1800
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000C&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000C&
      Height          =   2895
      Left            =   3840
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'Variables are not implemented in this release, but the functionality could
'easily be added by using Scripting.Dictionary and a SetVariable command.
'If you would like some help, e-mail me at Doobydog1@aol.com. I bet I can
'point you in the right direction.
'*** If you plan to use this code, make sure you vote for it ***!
Dim UserCodeArray() As String
Dim ArgumentsArray() As String
Dim UserCode As String
Dim LineFunction As String
Dim LineArguments As String
Dim UserLines As Integer
Dim LeftParenthesisPos As Integer
Dim RightParenthesisPos  As Integer
Dim LineLength As Integer
Dim i As Integer

'Clears debug textbox, delete it unless you want to view the debug info
Text2.Text = ""

'Preparing the code
UserCode = BackslashEscape(Text1.Text)

'Splitting the code, an array element for each command and its arguments
UserCodeArray() = Split(UserCode, ";")

'Finding number of commands
UserLines = UBound(UserCodeArray)

ReDim Preserve UserCodeArray(UserLines - 1)
For i = 0 To UserLines - 1
    'The guts of the parsing/processing routine, pretty self-explanatory
    LineLength = Len(UserCodeArray(i))
    LeftParenthesisPos = InStr(UserCodeArray(i), "(")
    RightParenthesisPos = InStrRev(UserCodeArray(i), ")")
    LineFunction = Left(UserCodeArray(i), LeftParenthesisPos - 1)
    LineArguments = Mid(UserCodeArray(i), LeftParenthesisPos + 1, RightParenthesisPos - (LeftParenthesisPos + 1))
        
        'THE COMMAND SELECT-CASE BLOCK
        'Each command has it's own case statement
        'Arguments are accessible through LineArguments string
        Select Case LineFunction
        
        Case "ExampleCommand"
            ArgumentsArray = Split(LineArguments, ",")
            For n = LBound(ArgumentsArray) To UBound(ArgumentsArray)
                ExampleCommand = ExampleCommand & ConvertEscapeCharsBack(ArgumentsArray(n))
            Next n
            MsgBox ExampleCommand
            
        'Case Add more commands here
        End Select
        
    'Displays debug info, not necessary to run, delete if you'd like
    Text2.Text = Text2.Text & DebugInfo(LineFunction, LineArguments)
Next
End Sub
Public Function BackslashEscape(Code As String) As String
'This function is also a good place to kill all tabs and newlines before
'we actually start processing the code
buffer = Replace(Code, vbCrLf, "")
buffer = Replace(buffer, vbTab, "")
'Replace all backslash escape characters so that we can process
'the agruments and convert them back later on
buffer = Replace(buffer, "\n", Chr(0) & "Newline")
buffer = Replace(buffer, "\t", Chr(0) & "Tab")
buffer = Replace(buffer, "\\", Chr(0) & "Backslash")
buffer = Replace(buffer, "\""", Chr(0) & "Quote")
buffer = Replace(buffer, "\;", Chr(0) & "Colon")
BackslashEscape = Replace(buffer, "\,", Chr(0) & "Comma")
End Function
Public Function ConvertEscapeCharsBack(Code As String) As String
'Convert the "intermediate" escape chars to the actual characters
buffer = Replace(Code, Chr(0) & "Newline", vbCrLf)
buffer = Replace(buffer, Chr(0) & "Tab", vbTab)
buffer = Replace(buffer, Chr(0) & "Backslash", "\")
buffer = Replace(buffer, Chr(0) & "Quote", """")
buffer = Replace(buffer, Chr(0) & "Colon", ";")
buffer = Replace(buffer, Chr(0) & "Comma", ",")
ConvertEscapeCharsBack = Trim(buffer)
End Function
Public Function DebugInfo(LineFunction As String, LineArguments As String) As String
Dim ArgumentArray() As String
DebugInfo = DebugInfo & "Command: " & LineFunction & vbCrLf
ArgumentArray = Split(LineArguments, ",")
For o = LBound(ArgumentArray) To UBound(ArgumentArray)
    TempArg = ConvertEscapeCharsBack(ArgumentArray(o))
    If Left(TempArg, 1) <> """" And Right(TempArg, 1) <> "" Then
        TempArg = """" & TempArg & """"
    End If
    DebugInfo = DebugInfo & "  " & "Argument: " & TempArg & vbCrLf
Next o
End Function


Private Sub Form_Load()

End Sub
