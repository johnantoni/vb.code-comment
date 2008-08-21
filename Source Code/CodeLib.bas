Attribute VB_Name = "CodeLibrary"
Option Explicit
'   Written by John A. Griffiths
'   (c) copyright JAG  2000
'   ~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'
'   Handles:
'       .bas    vb4-6 module files (text)
'       .gbl    vb3 module files (text)
'       .cls    class files (text)
'       .ctl    user-control files (text)
'       .frm    form code files (text)
'       .frx    form object files (binary, holds bitmaps, etc. used in form)
'       .mak    vb3 project files (text)
'       .res    32-bit resource files (binary)
'       .vbg    vb4-6 group project files (text)
'       .vbp    vb4-6 project files (text)
'
'
'       Private, Public, Friend or Standard Function Definitions
'       Private, Public, Friend or Standard Sub-Procedure Definitions
'       Private, Public, Friend or Standard Property Definitions
'
'       Fuinctional Definitions on multiple lines
'           (saves them as one single line)
'
'       Can also handle classes
'           (as they are defined like sub procedures)
'
'       produces list of variables not used at end of each procedure (code-block)
'           (rebuilds list after each successive run, does not damage comment header)
'


    'global constants
    '~~~~~~~~~~~~~~~~
    Public Const cNULL = "" 'empty string
    Public Const cSEP_URLDIR = "/" 'separator for dividing directories in URL addresses.
    Public Const cSEP_DIR = "\" 'directory separator character
    Public Const cBYVAL = "ByVal "
    Public Const cBYREF = "ByRef "
    Public Const cAS = " As "
    Public Const cEQUALS = " = "
    Public Const cSPC = "'   "
    Public Const cCOMMA = ", "
    Public Const cDIM = "Dim "
    Public Const cCONST = "Const "

    'used for scanning files for particular entries
    Public Const cFORM = "Form"
    Public Const cMOD = "Module"
    Public Const cCLASS = "Class"
    Public Const cUSERCONTROL = "UserControl"
    Public Const cRESOURCE = "ResFile32"

    'code definition names
    Public Const cSUB = "Sub"
    Public Const cFUNC = "Function"
    Public Const cPROP = "Property"

    'property types
    Public Const cPGET = "Property Get"
    Public Const cPSET = "Property Set"
    Public Const cPLET = "Property Let"

    'code definition types
    Public Const cPRIVATE = "Private "
    Public Const cPUBLIC = "Public "
    Public Const cFRIEND = "Friend "

    'commenting text (header & footer)
    Public Const cCOMMENTS = "'#-------------------------------------------------------------#"

    'date format constant
    Public Const cDFORMAT = "dd-mmm-yyyy"


    'global variables
    '~~~~~~~~~~~~~~~~
    Public gsFtype As String
    Public gsChar As String
    Public giLeftBit As Integer
    Public giRightBit As Integer
    Public giCounter As Integer
    Public gsOldPath As String
    Public gsNewPath As String
    Public garStyles() As CODESTYLE 'code-style storage array


    'type declarations
    '~~~~~~~~~~~~~~~~~
    Type VARTYPE 'for dead-variable checking
        Name As String
        Type As String
        Used As Boolean
    End Type

    Type CODESTYLE 'for storing code-style types
        Name As String
        Author As String
        Created As String
        History As String
    End Type

Public Function AddDIRSep(sPathName As String) As String
'#-------------------------------------------------------------#
' AUTHOR:       John Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   adds trailing directory path separator (back slash)
'   to end of path, unless one already exists
'
' PARAMETERS:
'   sPathName string:
'
' RETURNS:
'   String
'
' CHANGE HISTORY:
'   08-Mar-2000  John Griffiths (added front-end)
'#-------------------------------------------------------------#
    Dim sS As String

    sS = Trim(sPathName)
    If Right(Trim(sPathName), Len("\")) <> "\" Then sS = Trim(sPathName) & "\"

    AddDIRSep = sS

    'OLD CODE
    'If Right(Trim(sPathName), Len(cSEP_URLDIR)) <> cSEP_URLDIR And _
    '   Right(Trim(sPathName), Len(cSEP_DIR)) <> cSEP_DIR Then
    '    sPathName = RTrim(sPathName) & cSEP_DIR
    'End If
End Function
Public Sub CheckFileList()
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   work out which project in lstProjects to select, then call to
'   build a list of files that project uses
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    If (frmMain.lstProjects.ListCount = 0) Then
        'if lstProjects is empty, clear lstFiles and disable cmdAddComments
        frmMain.lstFiles.Clear
        frmMain.btnAddComments.Enabled = False

    Else
        'otherwise select first project in list and process it
        frmMain.lstProjects.Selected(0) = True
        MakeFileList (frmMain.lstProjects.List(0))
    End If
End Sub

Public Function DirExists(ByVal sDirName As String) As Integer
'#-------------------------------------------------------------#
' AUTHOR:       John Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   Checks if directory 'sDirName' exists
'   (can be used to check whether a floppy disk is in drive A: by passing "A:\")
'
'   //WORKING VERSION//
'
' PARAMETERS:
'   sDirName string:
'
' RETURNS:
'   Integer (exists = TRUE, doesn't exist = FALSE)
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    Dim sS As String
    Dim iMod As Integer

    On Error Resume Next

    sS = Dir(AddDIRSep(sDirName) & "*.*", vbDirectory) 'get directory name only
    iMod = Not (Len(sS) < 1)

    'modify for checkboxes
    If (iMod = -1) Then DirExists = 1 Else DirExists = 0

    Err = 0 'clear error flag
End Function

Public Sub MakeFileList(ByVal sFileName As String)
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      18-May-1999
' DESCRIPTION:
'   read project file and add it's component files to the listbox
'
' PARAMETERS:
'   sFileName string:
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'#-------------------------------------------------------------#
    Dim sMyLine As String
    Dim sINIValue As String
    Dim lMarker As Long
    Dim ilength As Integer
    Dim hfree As Integer


    'update input path textbox
    frmMain.txtFilePath.Text = frmMain.lstDirectories.Path & "\" & sFileName

    'clear lstFiles listbox
    frmMain.lstFiles.Clear

    'scan txtFilePath file for .frm and .bas entries (add to lstFiles)
    hfree = FreeFile 'assign free file-handle to hfree

    Open RTrim(frmMain.txtFilePath.Text) For Input As #hfree

    Do While Not EOF(hfree)
        Line Input #hfree, sMyLine 'put line into sMyLine
        sMyLine = Trim(sMyLine) 'trim line

        'search for FORM files
        If UCase(Left(sMyLine, Len(cFORM))) = UCase(cFORM) Then
            lMarker = InStr(sMyLine, "=")
            sINIValue = LTrim(Right(sMyLine, Len(sMyLine) - lMarker))
            frmMain.lstFiles.AddItem (sINIValue)

        'search for MODULE files
        ElseIf UCase(Left(sMyLine, Len(cMOD))) = UCase(cMOD) Then
            lMarker = InStr(sMyLine, ";")
            sINIValue = LTrim(Right(sMyLine, Len(sMyLine) - lMarker))
            frmMain.lstFiles.AddItem (sINIValue)

        'search for CLASS files
        ElseIf UCase(Left(sMyLine, Len(cCLASS))) = UCase(cCLASS) Then
            lMarker = InStr(sMyLine, ";")
            sINIValue = LTrim(Right(sMyLine, Len(sMyLine) - lMarker))
            frmMain.lstFiles.AddItem (sINIValue)

        'search for USERCONTROL files
        ElseIf UCase(Left(sMyLine, Len(cUSERCONTROL))) = UCase(cUSERCONTROL) Then
            lMarker = InStr(sMyLine, "=")
            sINIValue = LTrim(Right(sMyLine, Len(sMyLine) - lMarker))
            frmMain.lstFiles.AddItem (sINIValue)

        'search for 32-bit Resource files
        ElseIf UCase(Left(sMyLine, Len(cRESOURCE))) = UCase(cRESOURCE) Then
            lMarker = InStr(sMyLine, "=") + 1 'to account for the '"' at either end
            sINIValue = LTrim(Right(sMyLine, Len(sMyLine) - lMarker))
            sINIValue = Left(sINIValue, Len(sINIValue) - 1)
            frmMain.lstFiles.AddItem (sINIValue)
        End If
    Loop

    Close #hfree 'close file


    'find projects full path (not including proj file name)
    ilength = Len(Trim(frmMain.txtFilePath.Text))
    giCounter = 0

    While ((Not (giCounter = 2)) And (ilength > 0))
        gsChar = Mid(Trim(frmMain.txtFilePath.Text), ilength, 1)

        If (gsChar = "\") And (giCounter = 0) Then
            giRightBit = ilength - 1
            giCounter = 1
            gsChar = ""

        ElseIf (gsChar = "\") And (giCounter = 1) Then
            giLeftBit = ilength + 1
            giCounter = 2
            gsChar = ""
        End If

        ilength = ilength - 1
    Wend

    'put project's directory name in gsNewPath
    gsNewPath = Mid(Trim(frmMain.txtFilePath.Text), giLeftBit, giRightBit - giLeftBit + 1)

    'make gsOldPath point to original directory name (using gsNewPath)
    gsOldPath = Mid(Trim(frmMain.txtFilePath.Text), 1, giLeftBit - 2) & "\" & gsNewPath

    'make lstDirectories box point to txtfilepath directory (using gsNewPath)
    frmMain.lstDirectories.Path = Mid(Trim(frmMain.txtFilePath.Text), 1, giLeftBit - 2) & "\" & gsNewPath

    'set txtoutputdir to full path of project's dump directory
    frmMain.txtOutputDir.Text = Mid(Trim(frmMain.txtFilePath.Text), 1, giLeftBit - 2) & "\_" & gsNewPath

    'if nothing in lstfiles listbox, disable ADD COMMENTS button
    If (frmMain.lstFiles.ListCount = 0) Then
        frmMain.btnAddComments.Enabled = False
    Else
        frmMain.btnAddComments.Enabled = True
    End If

    'work out number of files (adjust text accordingly)
    If (frmMain.lstFiles.ListCount = 1) Then sMyLine = " file" Else sMyLine = " files"
    frmMain.lblFiles.Caption = "Files used in project (" & frmMain.lstFiles.ListCount & sMyLine & ") :"
End Sub

Public Sub WriteComments(ByVal iFileNum As Integer, ByVal sDecLine As String)
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'   comments template, and passed/returned code
'
' PARAMETERS:
'   iFileNum integer:
'   sDecLine integer:
'
' CHANGE HISTORY:
'   12-Nov-1998  'rtrim' bug fix
'   25-Nov-1998  pass function as parameter bug fix
'#-------------------------------------------------------------#
    Dim iStart As Integer
    Dim iEnd As Integer
    Dim iMarker1 As Integer
    Dim iMarker2 As Integer
    Dim iRetMarker As Integer
    Dim sVar As String
    Dim sName As String
    Dim sType As String
    Dim sReturns As String
    Dim bLineEnd As Boolean
    Dim bNoParams As Boolean
    Dim bFuncNone As Boolean
    Dim iFuncChecking As String
    Dim sTempStr As String
    Dim iTempLen As Integer
    Dim iTempCnt As Integer
    Dim iTempMarker As Integer


    sDecLine = Trim(sDecLine) 'trim our copy of line we are going to use

    Print #iFileNum, cCOMMENTS
    Print #iFileNum, "' AUTHOR:   " & Chr(9) & Trim(frmMain.txtAuthor.Text)
    Print #iFileNum, "' CREATED:  " & Chr(9) & Trim(frmMain.txtCreated.Text)
    Print #iFileNum, "' DESCRIPTION:" & Chr(9)

    'fill in description for well-known items
    If InStr(1, UCase(sDecLine), UCase("Sub Form_Load")) Then
        Print #iFileNum, cSPC & "Form Initialisation Code"
    ElseIf InStr(1, UCase(sDecLine), UCase("Sub Form_Unload")) Then
        Print #iFileNum, cSPC & "Form Unloading Code"
    Else
        Print #iFileNum, cSPC 'otherwise leave blank
    End If


    Print #iFileNum, "'"
    Print #iFileNum, "' PARAMETERS:"

    'initialise flags
    bLineEnd = False 'have we reached the end of the line
    bFuncNone = True 'for working out if it's a function with no parameters
    bNoParams = True

    'PARAMETERS handling code
    '++++++++++++++++++++++++
    'print parameter information
    iStart = InStr(1, sDecLine, "(")
    iMarker1 = iStart

    If Mid(sDecLine, Len(sDecLine), 1) = ")" Then
        'no returns
        iEnd = Len(sDecLine)
    Else
        'find ")" (going backwards along string "sDecLine")
        iTempLen = Len(Trim(sDecLine))

        'initialise temporary markers
        iTempCnt = 0: iTempMarker = 0

        While ((Not (iTempCnt = 1)) And (iTempLen > 0))
            sTempStr = Mid(sDecLine, iTempLen, 1)
            If (sTempStr = ")") Then
                iTempMarker = iTempLen
                iTempCnt = 1
                sTempStr = ""
            End If
            iTempLen = iTempLen - 1
        Wend

        iEnd = iTempLen + 1

        If (iStart + 1) = iEnd Then
            bFuncNone = True 'if function then there is no space for any parameters
        End If
    End If


    Do While (bLineEnd = False)
        iMarker2 = InStr(iMarker1, sDecLine, cCOMMA) 'was cAS

        If (iMarker2 = 0) Then
            If (iMarker1 < iEnd) And (Not ((iMarker1 + 1) = iEnd)) Then
                'reached partial end of declaration, keep on processing one more time
                iMarker2 = iEnd
            Else
                'reached end of declaration, stop processing
                iMarker1 = iTempMarker
                bLineEnd = True
            End If
        End If


        If ((Not (iMarker2 = 0)) And bLineEnd = False) Then
            'take out variable name, put in sVar
            sVar = Trim(Mid(sDecLine, iMarker1, iMarker2 - iMarker1))

            'if it's a function with no parameters, don't put them in
            '(by this time bLineEnd is set to true, so it will quit the loop afterwards)
            If (Not (iMarker1 > iMarker2)) Then 'was (iMarker2 > iMarker1)

                'if find "(" at start, remove it (a bug)
                If Left(sVar, 1) = "(" Then
                    sVar = Right(sVar, Len(sVar) - 1)
                End If


                If InStr(1, sVar, cBYVAL) > 0 Then
                    'remove BYVAL text from "sVar" before processing
                    sVar = Trim(Right(sVar, Len(sVar) - Len(cBYVAL)))

                ElseIf InStr(1, sVar, cBYREF) > 0 Then
                    'remove BYREF text from "sVar" before processing
                    sVar = Trim(Right(sVar, Len(sVar) - Len(cBYREF)))
                End If


                iTempCnt = InStr(1, sVar, cAS)
                If iTempCnt = 0 Then
                    'have not found " As ", so variable undefined
                    sType = "Variant"
                    sName = sVar
                Else
                    'found type, put in sType
                    sType = Right(sVar, (Len(sVar) - iTempCnt) - 3) '-3 to remove "As "
                    sName = Mid(sVar, 1, iTempCnt)
                End If

                'save variable details in comments template
                Print #iFileNum, cSPC & Trim(sName) & " " & LCase(sType) & ":"
                bFuncNone = False 'if function then has got parameters
            End If

            iMarker1 = iMarker2 + Len(cCOMMA) 'pass over values (advance to next position)
            bNoParams = False 'got parameters (function or sub)
        End If
    Loop

    'now deal with a no-parameters situation
    If (bNoParams = True) Then
        'if no parameters, say so
        Print #iFileNum, cSPC & "None"
    ElseIf ((InStr(sDecLine, cFUNC) <> 0) Or (InStr(sDecLine, cPGET) <> 0)) _
            And (bFuncNone = True) Then
        'if function and no parameters, say so
        Print #iFileNum, cSPC & "None"
    End If


    'RETURNS VALUE handling code
    '+++++++++++++++++++++
    iFuncChecking = 0 'initialise function testing flag

    'if function, add "Returns" section to "comments" text
    iFuncChecking = InStr(sDecLine, cFUNC)
    If iFuncChecking = 0 Then
        'if got nothing, test for Property Get's as well
        iFuncChecking = InStr(sDecLine, cPGET)
    End If

    If (iFuncChecking <> 0) Then
        If (bLineEnd = True) Then
            Print #iFileNum, "'"
            Print #iFileNum, "' RETURNS:"

            'does last character of code header contain ")"
            If (Right(Trim(sDecLine), 1) = ")") Then
                'check to see if array delcaration, if so build Returns Information
                If Right(Trim(sDecLine), 2) = "()" Then
                    GoSub ProcessReturns
                Else
                    'if so (and not array declaration)
                    'then do not return any Returns Info (even if function)
                    Print #iFileNum, cSPC & "None"
                End If

            Else
                'otherwise if ")" not found entirely, build Returns Information
                GoSub ProcessReturns
            End If
        End If
    End If

    Print #iFileNum, "'"
    Print #iFileNum, "' CHANGE HISTORY:"
    Print #iFileNum, "'   " & frmMain.txtHistory.Text
    Print #iFileNum, cCOMMENTS


EndHeader:
    Exit Sub


ProcessReturns:
    sReturns = ""
    iRetMarker = InStr(iTempMarker + 1, sDecLine, cAS) '+ 1 to account for the ")"

    'if got something, write it's type to the output file
    If iRetMarker <> 0 Then
        iRetMarker = iRetMarker + Len(cAS)
        sReturns = Trim(Mid(sDecLine, iRetMarker, Len(sDecLine)))
        Print #iFileNum, cSPC & sReturns 'write return type
    Else
        Print #iFileNum, cSPC & "None" 'write nothing
    End If
    Return 'go back to where jumped out
End Sub

Public Function CheckFinished(ByVal sCode As String) As Boolean
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'
'
' PARAMETERS:
'   sCode string:
'
' RETURNS:
'   Boolean
'
' CHANGE HISTORY:
'   08-Mar-2000  John A. Griffiths (code update)
'#-------------------------------------------------------------#
    If InStr(1, sCode, "End " & cSUB, vbTextCompare) Or _
       InStr(1, sCode, "End " & cFUNC, vbTextCompare) Or _
       InStr(1, sCode, "End " & cPROP, vbTextCompare) Then
        CheckFinished = True 'return true, reached end of code block

    Else
        CheckFinished = False 'return false, not reached end of code block
    End If
End Function

Public Sub WriteEndNote(ByRef iFileNum As Integer, ByVal sEndLine As String, ByRef aVarArray() As VARTYPE, ByVal iVarCounter As Integer)
'#-------------------------------------------------------------#
' AUTHOR:       John A. Griffiths
' CREATED:      4-Nov-1998
' DESCRIPTION:
'
'
' PARAMETERS:
'   iFileNum integer:
'   sEndLine string:
'   aVarArray() vartype:
'   iVarCounter integer:
'
' CHANGE HISTORY:
'   08-Mar-2000  John A. Griffiths (code update)
'#-------------------------------------------------------------#
    Dim iCnt As Integer
    Dim bList As Boolean


    If (iVarCounter > 0) Then
        bList = False 'initialise checking flag
        For iCnt = 1 To iVarCounter
            'if any variables are not used, set bList to TRUE
            If aVarArray(iCnt).Used = False Then bList = True
        Next

        'if any variables were not used, print NOT USED text at end of procedure
        If bList = True Then
            Print #iFileNum, ""
            Print #iFileNum, cCOMMENTS
            Print #iFileNum, "' VARIABLES NOT USED:"

            'for every variable not used, write a line to the source code file
            For iCnt = 1 To iVarCounter
                If aVarArray(iCnt).Used = False Then
                    Print #iFileNum, cSPC & aVarArray(iCnt).Name & " as " & LCase(aVarArray(iCnt).Type) '& ""
                End If
            Next
            Print #iFileNum, cCOMMENTS
        End If
    End If

    'print last line to file (usually "End Sub")
    Print #iFileNum, sEndLine
End Sub

Sub InitStyles()
'#-------------------------------------------------------------#
' AUTHOR:   	John Griffiths
' CREATED:  	04-Jun-2000
' DESCRIPTION:	
'   
'
' PARAMETERS:
'   None
'
' CHANGE HISTORY:
'   04-Jun-2000  John Griffiths (code update)
'#-------------------------------------------------------------#
    Dim lI As Long


    AddStyle "Default Style", "", Format(Date, cDFORMAT), Format(Date, cDFORMAT) & "  "
    AddStyle "JAG - code update", "John Griffiths", Format(Date, cDFORMAT), Format(Date, cDFORMAT) & "  John Griffiths (code update)"
    AddStyle "JAG - new program style", "John Griffiths", Format(Date, cDFORMAT), Format(Date, cDFORMAT) & "  John Griffiths (initial release)"

    For lI = LBound(garStyles) To UBound(garStyles)
        frmMain.lstStyle.AddItem garStyles(lI).Name
    Next
End Sub

Sub AddStyle(sName As String, sAuthor As String, sCreated As String, sHistory As String)
'#-------------------------------------------------------------#
' AUTHOR:   	John Griffiths
' CREATED:  	04-Jun-2000
' DESCRIPTION:	
'   
'
' PARAMETERS:
'   sName string:
'   sAuthor string:
'   sCreated string:
'   sHistory string:
'
' CHANGE HISTORY:
'   04-Jun-2000  John Griffiths (code update)
'#-------------------------------------------------------------#
    Dim lLength As Long
    Dim lCell As Long
    Dim bNoSize As Boolean

    On Error GoTo ArrayEmpty


    bNoSize = False

    lLength = UBound(garStyles)

    If (bNoSize = False) Then
        'increase array size by one, preserving previous cells
        ReDim Preserve garStyles(lLength + 1)
    End If

    lCell = UBound(garStyles)

    garStyles(lCell).Name = Trim(sName)
    garStyles(lCell).Author = Trim(sAuthor)
    garStyles(lCell).Created = Trim(sCreated)
    garStyles(lCell).History = Trim(sHistory)

    Exit Sub


ArrayEmpty:
    ReDim garStyles(0) 'initialise array by one
    bNoSize = True 'don't resize after this

    Resume Next
End Sub
