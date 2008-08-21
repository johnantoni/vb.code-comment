VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB Commentary"
   ClientHeight    =   6600
   ClientLeft      =   540
   ClientTop       =   1620
   ClientWidth     =   10440
   Icon            =   "comments.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6600
   ScaleWidth      =   10440
   Begin VB.Frame fmeStyle
      Caption         =   "Comment Details"
      ForeColor       =   &H8000000D&
      Height          =   3645
      Left            =   6330
      TabIndex        =   15
      Top             =   1020
      Width           =   3975
      Begin VB.ListBox lstStyle
         Height          =   840
         Left            =   510
         TabIndex        =   23
         Top             =   2640
         Width           =   3375
      End
      Begin VB.TextBox txtHistory
         Height          =   1065
         Left            =   90
         MultiLine       =   -1  'True
         TabIndex        =   18
         Top             =   1440
         Width           =   3795
      End
      Begin VB.TextBox txtCreated
         Height          =   285
         Left            =   750
         TabIndex        =   17
         Top             =   840
         Width           =   3135
      End
      Begin VB.TextBox txtAuthor
         Height          =   285
         Left            =   90
         TabIndex        =   16
         Top             =   450
         Width           =   3795
      End
      Begin VB.Label lblStyle
         BackStyle       =   0  'Transparent
         Caption         =   "Style"
         Height          =   210
         Left            =   90
         TabIndex        =   24
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label lblAuthor
         Caption         =   "Author"
         Height          =   210
         Left            =   90
         TabIndex        =   21
         Top             =   230
         Width           =   600
      End
      Begin VB.Label lblCreationDate
         Caption         =   "Created"
         Height          =   210
         Left            =   90
         TabIndex        =   20
         Top             =   890
         Width           =   600
      End
      Begin VB.Label lblChangeHistory
         Caption         =   "Change History"
         Height          =   210
         Left            =   90
         TabIndex        =   19
         Top             =   1205
         Width           =   2220
      End
   End
   Begin VB.Frame fmeProgress
      Caption         =   "Progress"
      ForeColor       =   &H8000000D&
      Height          =   825
      Left            =   6330
      TabIndex        =   12
      Top             =   4800
      Width           =   3975
      Begin ComctlLib.ProgressBar barProgress
         Height          =   285
         Left            =   90
         TabIndex        =   14
         Top             =   450
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   503
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblProgressMessage
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3645
      End
   End
   Begin VB.DirListBox lstDirectories
      Height          =   3015
      Left            =   120
      TabIndex        =   5
      Top             =   1470
      Width           =   3375
   End
   Begin VB.TextBox txtOutputDir
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   510
      Width           =   9165
   End
   Begin VB.FileListBox lstProjects
      Height          =   1455
      Left            =   120
      Pattern         =   "*.vbp;*.mak;*.vbg"
      TabIndex        =   6
      Top             =   4950
      Width           =   3375
   End
   Begin VB.DriveListBox lstDrives
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   1050
      Width           =   3375
   End
   Begin VB.CommandButton btnAddComments
      BackColor       =   &H80000016&
      Caption         =   "Add Comments"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6420
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6030
      Width           =   1695
   End
   Begin VB.TextBox txtFilePath
      Height          =   285
      Left            =   1140
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   9165
   End
   Begin VB.CommandButton btnExit
      Caption         =   "Exit"
      BeginProperty Font
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6030
      Width           =   1695
   End
   Begin VB.ListBox lstFiles
      Height          =   4935
      ItemData        =   "comments.frx":0442
      Left            =   3660
      List            =   "comments.frx":0444
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1470
      Width           =   2535
   End
   Begin VB.Label lblMessage
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Won't alter files already in 'Save To' directory"
      Height          =   255
      Left            =   6360
      TabIndex        =   22
      Top             =   5700
      Width           =   3945
   End
   Begin VB.Label lblProjects
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Project Files Found"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4740
      Width           =   3495
   End
   Begin VB.Label lblSavePath
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Save To"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   30
      TabIndex        =   10
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label lblProjectPath
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Project File"
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   30
      TabIndex        =   9
      Top             =   165
      Width           =   1035
   End
   Begin VB.Label lblFiles
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Files used in project :"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   3690
      TabIndex        =   8
      Top             =   1230
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAddComments_Click()
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   Code to add comments to VB project files (forms & modules).
'   Acting on 'ADD COMMENTS' button.
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    Dim aVars() As VARTYPE 'used to store variable details (not globals)
    Dim iVarCnt As Integer 'for keeping track of variables in array
    Dim sDumpDir As String
    Dim sFrxPath As String
    Dim sMyLine As String
    Dim sInFile As String
    Dim sOutFile As String
    Dim iInFile As Integer
    Dim iOutFile As Integer
    Dim bAddComments As Boolean
    Dim bVarChecking As Boolean
    Dim sTempLine As String
    Dim sNewRead As String
    Dim iArrLen As String
    Dim iCnt As Integer
    Dim iMarker1 As Integer
    Dim iMarker2 As Integer
    Dim iMarker3 As Integer
    Dim iMarker4 As Integer
    Dim bEndOfLine As Boolean
    Dim bEnding As Boolean


    iArrLen = 1000 'set array variable's length (may want to change later)

    bVarChecking = False
    btnAddComments.Enabled = False 'disable Add Comments button while processing


    'set sDumpDir to equal full path of project's new "dump" directory
    sDumpDir = Mid(Trim(txtFilePath.Text), 1, giLeftBit - 2) & "\_" & gsNewPath

    If (DirExists(sDumpDir) = False) Then
        MkDir sDumpDir 'create output directory, if not exist
    End If

    giCounter = 0
    barProgress.Max = lstFiles.ListCount
    barProgress.Value = giCounter


    'copy project files to new directory (overwrites existing files)
    For giCounter = 0 To (lstFiles.ListCount - 1)

        'now add comments template to files in lstFiles listbox
        lblProgressMessage.Caption = "Processing " & lstFiles.List(giCounter)
        sInFile = gsOldPath & "\" & lstFiles.List(giCounter)
        sOutFile = RTrim(txtOutputDir.Text & "\" & lstFiles.List(giCounter))


        If UCase(Right(lstFiles.List(giCounter), 4)) = UCase(".res") Then
            'if resource file, just copy it over (make no changes)
            FileCopy sInFile, sOutFile

        Else
            '(if outfile doesn't exist, add comments and save new file)
            If Len(Dir(sOutFile)) < 1 Then

                bAddComments = False 'initialise bAddComments flag

                iInFile = FreeFile
                Open sInFile For Input As #iInFile

                iOutFile = FreeFile
                Open sOutFile For Output As #iOutFile

                'scan file for procedures and functions
                '(add comments template where needed)
                Do While Not EOF(iInFile)
                    'read line and trim from right (e.g. don't remove previous indentations)
                    Line Input #iInFile, sMyLine 'put line into sMyLine
                    sMyLine = RTrim(sMyLine) 'trim from right

                    '(if function or sub, set flag to add comments)
                    If (Left(sMyLine, Len(cSUB)) = cSUB) Or _
                       (Left(sMyLine, Len(cPRIVATE & cSUB)) = (cPRIVATE & cSUB)) Or _
                       (Left(sMyLine, Len(cPUBLIC & cSUB)) = (cPUBLIC & cSUB)) Or _
                       (Left(sMyLine, Len(cFRIEND & cSUB)) = (cFRIEND & cSUB)) Or _
                       (Left(sMyLine, Len(cFUNC)) = cFUNC) Or _
                       (Left(sMyLine, Len(cPRIVATE & cFUNC)) = (cPRIVATE & cFUNC)) Or _
                       (Left(sMyLine, Len(cPUBLIC & cFUNC)) = (cPUBLIC & cFUNC)) Or _
                       (Left(sMyLine, Len(cFRIEND & cFUNC)) = (cFRIEND & cFUNC)) Or _
                       (Left(sMyLine, Len(cPGET)) = (cPGET)) Or _
                       (Left(sMyLine, Len(cPRIVATE & cPGET)) = (cPRIVATE & cPGET)) Or _
                       (Left(sMyLine, Len(cPUBLIC & cPGET)) = (cPUBLIC & cPGET)) Or _
                       (Left(sMyLine, Len(cFRIEND & cPGET)) = (cFRIEND & cPGET)) Or _
                       (Left(sMyLine, Len(cPSET)) = (cPSET)) Or _
                       (Left(sMyLine, Len(cPRIVATE & cPSET)) = (cPRIVATE & cPSET)) Or _
                       (Left(sMyLine, Len(cPUBLIC & cPSET)) = (cPUBLIC & cPSET)) Or _
                       (Left(sMyLine, Len(cFRIEND & cPSET)) = (cFRIEND & cPSET)) Or _
                       (Left(sMyLine, Len(cPLET)) = (cPLET)) Or _
                       (Left(sMyLine, Len(cPRIVATE & cPLET)) = (cPRIVATE & cPLET)) Or _
                       (Left(sMyLine, Len(cPUBLIC & cPLET)) = (cPUBLIC & cPLET)) Or _
                       (Left(sMyLine, Len(cFRIEND & cPLET)) = (cFRIEND & cPLET)) Then

                        'if encounter split-up code definitions join them together
                        '(will be saved as one long line, not split up)
                        Do While Right(sMyLine, 2) = " _"
                            sNewRead = ""
                            Line Input #iInFile, sNewRead
                            sNewRead = " " & Trim(sNewRead) 'trim fully and add space
                            sMyLine = Left(sMyLine, Len(sMyLine) - 2) & sNewRead
                        Loop

                        bAddComments = True
                        sTempLine = sMyLine 'for passing to writecomments()

                    '(if comments already there, dont add them)
                    ElseIf (bAddComments = True) Then
                        If (Not (Left(sMyLine, Len(cCOMMENTS)) = cCOMMENTS)) Then
                            WriteComments iOutFile, sTempLine
                        End If
                        bAddComments = False 'reset bAddComments flag

                        'setup variable storage array
                        ReDim aVars(iArrLen)

                        'initialise variable array
                        For iCnt = 0 To iArrLen
                            aVars(iCnt).Used = False
                        Next

                        iVarCnt = 0 'set variable counter to 0
                        bVarChecking = True 'begin dead-variable checking
                    End If


                    If (bVarChecking = True) Then
                        If CheckFinished(sMyLine) = True Then
                            'if reached end of code-block,
                            'print list of vars unused in array (if any)
                            WriteEndNote iOutFile, sMyLine, aVars, iVarCnt

                            'then reset var-checking variables
                            iVarCnt = 0 'reset variable counter
                            bVarChecking = False 'turn off checking flag
                            Erase aVars 'release memory used by array for other functions
                        Else
                            'otherwise build array of variables not used in code
                            GoSub BuildAndCheck
                            Print #iOutFile, sMyLine 'then print line normally
                        End If
                    Else
                        'copy source line to new .frm (or .bas) code file
                        Print #iOutFile, sMyLine
                    End If
                Loop

                'close opened files
                Close #iInFile
                Close #iOutFile
            End If

            'copy this form's related .frx file if .frm exists
            sFrxPath = Dir(gsOldPath & "\" & Left(lstFiles.List(giCounter), Len(lstFiles.List(giCounter)) - 4) & ".frx")
            lblProgressMessage.Caption = "Copying " & (Left(lstFiles.List(giCounter), Len(lstFiles.List(giCounter)) - 4) & ".frx")

            If (Not (sFrxPath = "")) Then
                FileCopy gsOldPath & "\" & sFrxPath, Trim(txtOutputDir.Text) & "\" & Left(lstFiles.List(giCounter), Len(lstFiles.List(giCounter)) - 4) & ".frx"
            End If
        End If

        barProgress.Value = giCounter
    Next

    'copy project file to output directory
    FileCopy txtFilePath.Text, sDumpDir & "\" & Mid((Trim(txtFilePath.Text)), giRightBit + 1, Len(Trim(txtFilePath.Text)) - giRightBit)

    btnAddComments.Enabled = True 'enable button once finished processing

    lblProgressMessage.Caption = "Task Complete"
    barProgress.Value = 0


ExitCode:
    Exit Sub


BuildAndCheck:
    If InStr(1, sMyLine, "ReDim ", vbBinaryCompare) = 0 Then
        If InStr(1, sMyLine, cDIM, vbBinaryCompare) > 0 Then
            'variable declaration found, parse line and add to aVars
            bEndOfLine = False
            bEnding = False
            iMarker1 = InStr(1, sMyLine, cDIM, vbBinaryCompare) + Len(cDIM) 'start of name
            iMarker4 = InStr(iMarker1, sMyLine, cCOMMA, vbBinaryCompare) 'end of type
            If iMarker4 = 0 Then iMarker4 = Len(sMyLine) + 1

            Do While (bEndOfLine = False)
                iVarCnt = iVarCnt + 1

                'if there is an " As " in the declaration
                If InStr(1, Mid(sMyLine, iMarker1, iMarker4 - iMarker1), cAS, vbBinaryCompare) > 0 Then
                    'type declared explicitly
                    iMarker2 = InStr(iMarker1, sMyLine, cAS, vbBinaryCompare) 'end of name
                    iMarker3 = iMarker2 + Len(cAS) 'start of type

                    'If InStr(1, Mid(sMyLine, iMarker1, iMarker2 - iMarker1), "()") Then
                    '    'found array declaration, remove brackets
                    '    aVars(iVarCnt).Name = Mid(sMyLine, iMarker1, (iMarker2 - iMarker1) - 2)
                    'Else
                        aVars(iVarCnt).Name = Mid(sMyLine, iMarker1, iMarker2 - iMarker1)

                        'deal with arrays better
                        If InStr(1, aVars(iVarCnt).Name, "(") > 0 Then
                            aVars(iVarCnt).Name = Mid(aVars(iVarCnt).Name, 1, (InStr(1, aVars(iVarCnt).Name, "(")) - 1)
                        End If
                    'End If

                    aVars(iVarCnt).Type = Mid(sMyLine, iMarker3, iMarker4 - iMarker3)

                Else
                    'type not declared explicitly
                    'If InStr(1, Mid(sMyLine, iMarker1, iMarker4 - iMarker1), "()") Then
                    '    'found array declaration, remove brackets
                    '    aVars(iVarCnt).Name = Mid(sMyLine, iMarker1, (iMarker2 - iMarker1) - 2)
                    'Else
                        aVars(iVarCnt).Name = Mid(sMyLine, iMarker1, iMarker4 - iMarker1)
                        'deal with arrays better
                        If InStr(1, aVars(iVarCnt).Name, "(") > 0 Then
                            aVars(iVarCnt).Name = Mid(aVars(iVarCnt).Name, 1, (InStr(1, aVars(iVarCnt).Name, "(")) - 1)
                        End If
                    'End If
                    aVars(iVarCnt).Type = "undeclared type"

                End If

                aVars(iVarCnt).Used = False 'initially set it to false (not used)


                iMarker1 = iMarker4 + Len(cCOMMA) 'start of name
                iMarker4 = InStr(iMarker1, sMyLine, cCOMMA, vbBinaryCompare) 'end of type

                If bEnding = True Then
                    bEndOfLine = True
                ElseIf iMarker1 > 0 And iMarker4 = 0 Then
                    If (Len(sMyLine) + 1) < iMarker1 Then
                        bEndOfLine = True
                    Else
                        iMarker4 = Len(sMyLine) + 1
                        bEnding = True
                    End If
                ElseIf iMarker1 = 0 Then
                    bEndOfLine = True
                End If
            Loop


        ElseIf InStr(1, sMyLine, cCONST, vbBinaryCompare) Then
            'constant declaration found, find constant's name
            iMarker1 = InStr(1, sMyLine, cCONST, vbTextCompare) + Len(cCONST)
            iMarker2 = InStr(iMarker1, sMyLine, cEQUALS, vbTextCompare)

            iVarCnt = iVarCnt + 1 'increment variable counter

            'record constant's details in variable array
            aVars(iVarCnt).Name = Trim(Mid(sMyLine, iMarker1, (iMarker2 - iMarker1)))
            aVars(iVarCnt).Type = Trim(cCONST)

        Else
            'otherwise check to see if code text matches any variables in aVars
            For iCnt = 1 To iVarCnt
                'do case sensitive search
                If aVars(iCnt).Used = False Then
                    If (sMyLine <> "") Then
                        If InStr(1, sMyLine, aVars(iCnt).Name, vbBinaryCompare) Then
                            'if found set that variable's USED flag to true
                            aVars(iCnt).Used = True
                        End If
                    End If
                End If
            Next
        End If
    End If

    Return 'return to calling code
End Sub

Private Sub btnExit_Click()
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   close program code
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    Unload frmMain 'close main form
End Sub

Private Sub Form_Activate()
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   08-Mar-2000  John A. Griffiths (code update)
'#-------------------------------------------------------------#
    frmMain.lstDirectories.Refresh 'update project files listbox
End Sub

Private Sub lstDirectories_Change()
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   change directory
'   (updates file list box to synchronize with directory list box)
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    Dim sMyTxt As String


    lblProgressMessage.Caption = ""
    lstProjects.Path = lstDirectories.Path

    'work out number of project files (adjust text accordingly)
    If frmMain.lstProjects.ListCount = 1 Then sMyTxt = " file" Else sMyTxt = " files"

    frmMain.lblProjects.Caption = "Project files found (" & Trim(LCase(lstProjects.Pattern)) & ") " & frmMain.lstProjects.ListCount & sMyTxt & " :"
    frmMain.lblFiles.Caption = "Files used in project (0 files) :"

    CheckFileList
End Sub

Private Sub lstDrives_Change()
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   change drive code
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    On Error GoTo DriveHandler


    lstDirectories.Path = lstDrives.Drive
    lblProgressMessage.Caption = ""


ExitHandler:
    Exit Sub


DriveHandler:
    lstDrives.Drive = lstDirectories.Path
    Exit Sub
End Sub

Private Sub lstProjects_Click()
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   select project file code
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    MakeFileList (lstProjects.List(lstProjects.ListIndex))
    lblProgressMessage.Caption = ""
End Sub

Private Sub Form_Load()
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   form initialisation code
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    'show version number
    frmMain.Caption = frmMain.Caption & " v" & App.Major & "." & App.Minor & "." & App.Revision

    'centre form on screen (avoiding windows start-bar at bottom)
    Move (Screen.Width - Width) / 2, (((Screen.Height - (Screen.TwipsPerPixelY * 26)) - Height) / 2)

    'initialise form
    lblChangeHistory.Caption = "Change History (" & Trim(LCase(cDFORMAT)) & ")"
    lblProgressMessage = ""
    barProgress.Value = 0

    lstFiles.Clear
    lstDirectories.Path = App.Path

    CheckFileList

    InitStyles

    lstStyle.Selected(0) = True
End Sub

Private Sub lstStyle_Click()
'#-------------------------------------------------------------#
' AUTHOR:       John .A. Griffiths
' CREATED:      21-May-1999
' DESCRIPTION:
'   change cemmenting style to use
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   21-May-1999
'#-------------------------------------------------------------#
    txtAuthor.Text = garStyles(lstStyle.ListIndex).Author
    txtCreated.Text = garStyles(lstStyle.ListIndex).Created
    txtHistory.Text = garStyles(lstStyle.ListIndex).History
End Sub
