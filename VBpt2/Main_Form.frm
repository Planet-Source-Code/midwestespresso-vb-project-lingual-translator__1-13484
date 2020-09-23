VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "VBpt2"
   ClientHeight    =   2625
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   8
      Text            =   "Save File As..."
      Top             =   960
      Width           =   4095
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   4800
      TabIndex        =   7
      Top             =   1560
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   765
      Left            =   1320
      TabIndex        =   3
      Top             =   1395
      Width           =   4095
   End
   Begin VB.ComboBox Combo2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "Main_Form.frx":0000
      Left            =   3480
      List            =   "Main_Form.frx":0019
      TabIndex        =   2
      Text            =   "Spanish"
      Top             =   480
      Width           =   1935
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      ItemData        =   "Main_Form.frx":005B
      Left            =   1320
      List            =   "Main_Form.frx":0074
      TabIndex        =   1
      Text            =   "English"
      Top             =   480
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   0
      Text            =   "Open File..."
      Top             =   60
      Width           =   4095
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   4800
      TabIndex        =   14
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Height          =   375
      Left            =   5160
      TabIndex        =   13
      Top             =   2280
      Width           =   375
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   12
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   480
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackColor       =   &H0052A6E4&
      Height          =   135
      Left            =   2280
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Reading / Writing Translated Form"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0052A6E4&
      Height          =   255
      Left            =   2160
      TabIndex        =   5
      Top             =   840
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   2625
      Left            =   0
      Picture         =   "Main_Form.frx":00B6
      Top             =   0
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()
'Step 1: Load File and Translate
    'Part A: Declarations
    Dim FileStr As String
    Dim FileN As String
    Dim TempStr As String
    Dim DotPos As Integer
    Dim tmpHolding
    Dim tmpDeclare
    Dim tmpFileName As String
    Dim refLang
    Dim refIndexLanguage As Integer
    Dim idxNumber As Integer
    Dim selFile As Integer
    'Part B: Check Languages
  
    If Combo1.Text = Combo2.Text Then
        Beep
        Exit Sub
    ElseIf Combo1.Text = "" Then
        Beep
        Exit Sub
    ElseIf Combo2.Text = "" Then
        Beep
        Exit Sub
    Else
    End If
      Text1.Visible = False
    Text2.Visible = False
    Text3.Visible = False
    Text4.Visible = False
    Combo1.Visible = False
    Combo2.Visible = False
    Label2.Visible = True
    Label3.Visible = True
    Label2.Caption = "Checking Languages"
    Label3.Width = Label3.Width * (1 / 6)
    'Part C: Load File
    On Err GoTo Go_Beep
    Label2.Caption = "Loading File"
    Label3.Width = Label3.Width * (2 / 6)
    If Text1.Text = "" Then Exit Sub
    Open Text1.Text For Input As #1
    FileStr = ""
    Do Until EOF(1)
    Line Input #1, TempStr
    FileStr = FileStr & TempStr & Chr$(13) & Chr$(10)
    Loop
    Text2.Text = FileStr
    Close #1
    'Part D: Replace Quotation Marks
    Label2.Caption = "Replacing Quotes"
    Label3.Width = Label3.Width * (3 / 6)
    Text2.Text = sReplaceCharacters(Text2.Text, Chr(34), " ~Q ")
    'Part E: Prepare Language Database
    Label2.Caption = "Preparing Database"
    Label3.Width = Label3.Width * (4 / 6)
    selFile = FreeFile
    tmpFileName = App.Path & "\RosettaStone\Dictionaries\" & Combo1.Text & "To" & Combo2.Text & ".txt"
    Open tmpFileName For Input As selFile
    refLang = Split(Text2, " ")
    Text3 = ""
    For refIndexLanguage = 1 To 3100
        Input #selFile, tmpDeclare
        tmpHolding = Split(tmpDeclare, vbTab)
        Trim (tmpHolding(0))
        For idxNumber = 0 To UBound(refLang)
        If LCase(refLang(idxNumber)) = LCase(tmpHolding(0)) Then
            refLang(idxNumber) = tmpHolding(1)
        End If
        Next idxNumber
    Next refIndexLanguage
    For idxNumber = 0 To UBound(refLang)
        Text3 = Text3 & " " & refLang(idxNumber)
        Text3 = Trim(Text3)
    Next idxNumber
    Text2.Text = Text3.Text
    Close selFile
    'Part F: Input Quotation Marks
    Label2.Caption = "Inputing Quotes"
    Label3.Width = Label3.Width * (5 / 6)
    Text2.Text = sReplaceCharacters(Text2.Text, " ~Q ", Chr(34))
    'Part G: Save the file
    Label2.Caption = "Saving"
    Label3.Width = Label3.Width
    If Text4.Text = "" Then Exit Sub
    Open Text4.Text For Output As #2
    Print #2, Text2
    Close #2
    Label2.Visible = False
    Label3.Visible = False
    Text1.Visible = True
    Text2.Visible = True
    Text4.Visible = True
    Combo1.Visible = True
    Combo2.Visible = True
Go_Beep:
Beep
End Sub

Private Sub Label4_Click()
End
End Sub

Private Sub Label5_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Combo1.Text = ""
Combo2.Text = ""

End Sub

Private Sub Label6_Click()
Text1.Text = "C:\"
End Sub

Private Sub Label7_Click()
Text1.Text = "Open File..."
Text2.Text = ""
Text3.Text = ""
Text4.Text = "Save Form As..."
Combo1.Text = "English"
Combo2.Text = "Spanish"
End Sub

Private Sub Label8_Click()
End
End Sub

Private Sub Label9_Click()
Form1.WindowState = 1
End Sub
