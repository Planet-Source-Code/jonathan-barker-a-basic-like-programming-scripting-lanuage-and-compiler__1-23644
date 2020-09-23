VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Jonny's Experimental Language..."
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "lang1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1680
      TabIndex        =   4
      Text            =   "prog.exe"
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1680
      TabIndex        =   3
      Text            =   "prog.txt"
      Top             =   360
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Make EXE"
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RunFile"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "LoadFile"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' A simple BASIC-style interpreter
'
'     -- By Jonny Barker
'
' This is an extremely simple language interpreter.
' it interprets a language similar to BASIC, but
' shows how to write interpreters, and could be
' adapted for any use at all, eg:
'       a scripting language/macros
'       a game language/mods
'          and anything you can think of...
'
' To add Functions:
' ~~~~~~~~~~~~~~~~~
'
' - In Form1_Load add the function name and increment
'   the funcs() array size.
' - Go down to dofunc and add your entry (see the
'   msgbox one for an example
' - That's it!
'
' Known Problems
' ~~~~~~~~~~~~~~
' - It   interprets    mathematical    expressions
'   consecutively, not in order of mathematical
'   precidence. (use lots of brackets!)
'   (If you want to correct this look up the 'railroad'
'   algorithm)
' - It is rather memory hungry (all those booleans)
'   (in small programs this should not matter)
' - Very little (useful) error handling...
' - Slow (too many Redims)
'
' To Add
' ~~~~~~
' - Proper mathematical expression handling
' - Proper boolean maths (a proper 'if' function)
' - Object orientation


Const WS_EX_STATICEDGE = &H20000
Const WS_EX_TRANSPARENT = &H20&
Const WS_CHILD = &H40000000
Const CW_USEDEFAULT = &H80000000
Const SW_NORMAL = 1
Private Type CREATESTRUCT
    lpCreateParams As Long
    hInstance As Long
    hMenu As Long
    hWndParent As Long
    cy As Long
    cx As Long
    Y As Long
    X As Long
    style As Long
    lpszName As String
    lpszClass As String
    ExStyle As Long
End Type
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Dim mWnd As Long
Dim tokens() As Token
Dim memory() As memoryentry
Dim forstack() As forstackentry
Dim forstackpos As Integer
Dim currenttoken As Integer
Dim memorytop As Integer
Dim funcs() As String
Dim i As Integer
Dim currentstack As Integer

Private Type memoryentry
    text As String
    info As Variant
End Type

Private Type forstackentry
    fortok As Integer
    varref As String
    nfrom As Variant
    nto As Variant
    nstep As Variant
End Type

Private Type funcargs
    funcname As String
    args() As Variant
    argrefs() As Variant 'for 'byrefs'
End Type
    
'yes, horrifically untidy, but that's life... (it's VB!!!)
'would use BitSets or bitvectors... (OR CLASSES!)
Private Type Token
    text As String 'text of the token
    pinttype As Integer 'NU
    assignment As Boolean 'is a =
    ident As Boolean 'is an identifier (variable name)
    func As Boolean 'is a function name
    openbracket As Boolean 'is a (
    delim As Boolean 'is a ,
    closebracket As Boolean 'is a )
    newline As Boolean 'is a \n
    stringlit As Boolean 'is a string literal ("hi")
    numlit As Boolean 'is a numerical literal (9)
    lit As Boolean 'literal expression
    mathop As Boolean 'is a mathematical operator
    add As Boolean 'is a +
    subtract As Boolean 'is a -
    divide As Boolean 'is a /
    multiply As Boolean 'is a *
    gotolabel As Boolean 'a goto label, eg "start:"
End Type
Private Sub Command1_Click()
ReDim tokens(0)
Dim flen As Long
currenttoken = 0
If InStr(1, Command, "debug") > 0 Then
    ChDir App.Path
    Open Text1 For Input As #1
Else
    Open App.Path & "\" & App.EXEName & ".exe" For Binary As #1
    flen = LOF(1)
    Dim strr As String
    strr = Space(8)
    Get #1, flen - 7, strr
    If strr <> "JSoftEXE" Then
        MsgBox "Invalid EXE!!!"
        End
    End If
    Dim ll As Long
    Get #1, flen - 11, ll
    'll is added flen
    Seek #1, flen - ll - 11
End If

Do
    Line Input #1, inline
    inline = Trim(inline)
backhere: 'aaah! nasty goto statements!
    'prelim whitespace removal
    isquote = False 'inside quote
    For i = 1 To Len(inline)
        tmpchar = Mid(inline, i, 1)
        If Asc(tmpchar) < 32 Then tmpchar = " "
        If tmpchar = " " And isquote = False Then
            inline = Left(inline, i - 1) + Mid(inline, i + 1)
            GoTo backhere
        ElseIf tmpchar = "'" And isquote = False Then
            'ignore the rest of the line
            inline = Left(inline, i - 1)
            Exit For
        ElseIf tmpchar = Chr(34) And isquote = False Then
            isquote = True
        ElseIf tmpchar = Chr(34) And isquote = True Then
            isquote = False
        End If
    Next i
    
'    Debug.Print inline
    
    isquote = False
    Dim tmpar As Token
        
    For i = 1 To Len(inline)
        tmpchar = Mid(inline, i, 1)
        If isdelim(tmpchar) = True And isquote = False Then
            'new token
            If tmpstr <> "" Then
                tmpar = getTokenType(tmpstr)
                addToken tmpar
            End If
            tmpar = getTokenType(tmpchar)
            addToken tmpar
            tmpstr = ""
        ElseIf tmpchar = Chr(34) And isquote = False Then
            If tmpstr <> "" Then
                tmpar = getTokenType(tmpstr)
                addToken tmpar
                tmpstr = ""
            End If
            isquote = True
        ElseIf tmpchar = Chr(34) And isquote = True Then
            isquote = False
            Dim tmpr As Token
            tmpr.text = tmpstr
            tmpr.stringlit = True
            tmpr.lit = True
            addToken tmpr
            tmpstr = ""
        Else
            tmpstr = tmpstr + tmpchar
        End If
        
        
    Next i

    If tmpstr <> "" Then
        Dim tmpy As Token
        tmpy = getTokenType(tmpstr)
        addToken tmpy
        tmpstr = ""
   End If

        Dim tmpa As Token
        tmpa.newline = True
        addToken tmpa

If Not InStr(1, Command, "debug") > 0 Then
    Debug.Print Seek(1)
    If Seek(1) = flen - 11 Then Exit Do
End If

Loop Until EOF(1)
Close #1

If InStr(1, Command, "debug") > 0 Then Beep


End Sub

Private Sub Command2_Click()

Dim tok As Token
Dim nexttok As Token
ReDim memory(0)
ReDim forstack(0)
i = 0
memorytop = 0

If InStr(1, Command, "debug") > 0 Then Form1.Hide

For i = 0 To currenttoken

tok = tokens(i)
If (i < currenttoken) Then nexttok = tokens(i + 1)

If tok.ident And nexttok.assignment Then
    'assume LET
    identname = tok.text
    i = i + 1
    tok = tokens(i)
    If (i < currenttoken) Then nexttok = tokens(i + 1)
        
    If nexttok.lit And tokens(i + 2).mathop Then
        i = i + 1
        setMemory identname, evalExpr
    ElseIf nexttok.lit Then
        setMemory identname, nexttok.text
    ElseIf nexttok.func Then
        i = i + 1
        setMemory identname, evalExpr
    ElseIf nexttok.openbracket Then
        i = i + 1
        setMemory identname, evalDelimExpr
    ElseIf nexttok.ident And tokens(i + 2).mathop Then
        i = i + 1
        setMemory identname, evalExpr
    ElseIf nexttok.ident Then
        setMemory identname, getMemory(nexttok.text)
    Else
        MsgBox "Expected Identifier!"
    End If
    
'    setMemory
ElseIf tok.func Then ' And (nexttok.openbracket) Then
    evalExpr
End If

DoEvents 'smash any hopes of getting even -some- speed in vb

Next i

If InStr(1, Command, "debug") > 0 Then Form1.Show
Form2.SetFocus

End Sub

Private Function evalDelimExpr()
Dim tok As Token
Dim nexttok As Token
Dim CS As Variant

    
tok = tokens(i)
If (i < currenttoken) Then nexttok = tokens(i + 1)
    
    If tok.openbracket Or (nexttok.mathop) Then 'tok.lit And
        'bracketed expression...
'        If tok.openbracket Then i = i + 1
        tok = tokens(i)
        If (i < currenttoken) Then nexttok = tokens(i + 1)
        
        If tok.numlit Then
            CS = Val(tok.text)
        ElseIf tok.func Then
            CS = evalExpr()
        ElseIf tok.openbracket Then
            CS = evalExpr()
        ElseIf tok.stringlit Then
            CS = tok.text
        ElseIf tok.ident Then
            CS = getMemory(tok.text)
        End If
                
        i = i + 1
        Dim s As Boolean
        
        Do
        tok = tokens(i)
        If (i < currenttoken) Then nexttok = tokens(i + 1)
        If tok.mathop Then
            If nexttok.numlit Then
                vr = Val(nexttok.text)
            ElseIf nexttok.openbracket Then
                i = i + 1
                vr = evalExpr()
                i = i - 1
                s = True
            ElseIf nexttok.func Then
                vr = evalExpr()
            ElseIf nexttok.stringlit Then
                vr = nexttok.text
            ElseIf nexttok.ident Then
                vr = getMemory(nexttok.text)
            Else
                MsgBox "Error! Invalid operator!"
            End If
                
            If Not IsNumeric(CS) Then
                'cs is text
                vr = CStr(vr)
            End If
            If Not IsNumeric(vr) Then
                'vr is text
                CS = CStr(CS)
            End If
            
            If tok.add Then CS = CS + vr
            If tok.subtract Then CS = Val(CS) - Val(vr)
            If tok.multiply Then CS = Val(CS) * Val(vr)
            If tok.divide Then CS = Val(CS) / Val(vr)
        
'            tok = tokens(i)
'            If (i < currenttoken) Then nexttok = tokens(i + 1)
        
        End If
        If (tok.closebracket Or tok.newline Or tok.delim) And s = False Then
            evalDelimExpr = CS
'            i = i + 1
            Exit Function
        End If
        s = False
        i = i + 2
        Loop
        
    End If



End Function

Private Function evalExpr()
Dim tok As Token
Dim nexttok As Token
Dim CS As Variant

    
tok = tokens(i)
If (i < currenttoken) Then nexttok = tokens(i + 1)
    
    If tok.openbracket Or (nexttok.mathop) Then 'tok.lit And
        'bracketed expression...
        If tok.openbracket Then i = i + 1
        tok = tokens(i)
        If (i < currenttoken) Then nexttok = tokens(i + 1)
        
        If tok.numlit Then
            CS = Val(tok.text)
        ElseIf tok.func Then
            CS = evalExpr()
        ElseIf tok.openbracket Then
            CS = evalExpr()
        ElseIf tok.stringlit Then
            CS = tok.text
        ElseIf tok.ident Then
            CS = getMemory(tok.text)
        End If
                
        i = i + 1
        Dim s As Boolean
        
        Do
        tok = tokens(i)
        If (i < currenttoken) Then nexttok = tokens(i + 1)
        If tok.mathop Then
            If nexttok.numlit Then
                vr = Val(nexttok.text)
            ElseIf nexttok.openbracket Then
                i = i + 1
                vr = evalExpr()
                i = i - 1
                s = True
            ElseIf nexttok.func Then
                vr = evalExpr()
            ElseIf nexttok.stringlit Then
                vr = nexttok.text
            ElseIf nexttok.ident Then
                vr = getMemory(nexttok.text)
            Else
                MsgBox "Error! Invalid operator!"
            End If
                
            If Not IsNumeric(CS) Then
                'cs is text
                vr = CStr(vr)
            End If
            If Not IsNumeric(vr) Then
                'vr is text
                CS = CStr(CS)
            End If
            
            If tok.add Then CS = CS + vr
            If tok.subtract Then CS = Val(CS) - Val(vr)
            If tok.multiply Then CS = Val(CS) * Val(vr)
            If tok.divide Then CS = Val(CS) / Val(vr)
        
'            tok = tokens(i)
'            If (i < currenttoken) Then nexttok = tokens(i + 1)
        
        End If
        If (tok.closebracket Or tok.newline Or tok.delim) And s = False Then
            evalExpr = CS
'            i = i + 1
            Exit Function
        End If
        s = False
        i = i + 2
        Loop
        
    End If
    
    'function... YAY
    Dim fa As funcargs
    ReDim fa.args(0)
    fa.funcname = tok.text
    currentarg = 0

    If nexttok.openbracket Then i = i + 2 Else i = i + 1
    tok = tokens(i)
    If (i < currenttoken) Then nexttok = tokens(i + 1)
        'tok should now be ident/lit or )

    Do Until tok.closebracket Or tok.newline
        currentarg = currentarg + 1
        ReDim Preserve fa.args(currentarg)
        ReDim Preserve fa.argrefs(currentarg)
        If nexttok.mathop Then
            fa.args(currentarg) = evalExpr()
            i = i - 1
            tok = tokens(i)
            If (i < currenttoken) Then nexttok = tokens(i + 1)
        ElseIf tok.numlit Then
            fa.args(currentarg) = Val(tok.text)
        ElseIf tok.stringlit Then
            fa.args(currentarg) = tok.text
        ElseIf tok.ident Then
            fa.args(currentarg) = getMemory(tok.text)
            fa.argrefs(currentarg) = tok.text
        ElseIf tok.openbracket Then
            fa.args(currentarg) = evalDelimExpr()
            If tokens(i).delim Then i = i - 1
            tok = tokens(i)
            If (i < currenttoken) Then nexttok = tokens(i + 1)
        Else
            MsgBox "Invalid identifier..."
        End If
        
        tok = tokens(i)
        If (i < currenttoken) Then nexttok = tokens(i + 1)
    
    If nexttok.closebracket Or nexttok.newline Then Exit Do
    
    If Not nexttok.delim Then
        MsgBox "Expected Delimiter or )!"
    End If
    
    i = i + 2
    tok = tokens(i)
    If (i < currenttoken) Then nexttok = tokens(i + 1)
    
    Loop
    
    evalExpr = dofunc(fa, currentarg)
    
End Function

Private Sub Command3_Click()
'If LCase(Right(App.EXEName, 4)) <> ".exe" Then
'    MsgBox "Do this after compiled! DOOF!"
'    Exit Sub
'End If

ChDir App.Path
Open Text2 For Binary As #1
Open Text1 For Binary As #2
Dim strr As String
strr = Space(8)
Get #1, LOF(1) - 8, strr
If strr = "JSoftEXE" Then
    MsgBox "Already DONE!!!"
    Exit Sub
End If
Seek #1, LOF(1)
strr = Space(LOF(2))
Get #2, , strr
Put #1, , strr
Dim l As Long
l = LOF(2)
Put #1, , l
Put #1, , "JSoftEXE"
Close #1
Close #2
MsgBox "Done!"
End Sub

' TYPES...
' 0  '=' assignment
' 1      ident
' 2      func
' 3  '(' open paranthese
' 4  ',' delimeter
' 5  ')' close paranthese
' 6      newline
' 7      string literal

Private Sub Form_Load()
ReDim Preserve funcs(13) ' num of highest funcs(n)
funcs(0) = "msgbox"
funcs(1) = "ifequal"
funcs(2) = "ifless"
funcs(3) = "ifgreater"
funcs(4) = "inputbox"
funcs(5) = "createwindow" ' needs writing
funcs(6) = "beep"
funcs(7) = "goto" ' heh NASTY!
funcs(8) = "for"
funcs(9) = "next"
funcs(10) = "input"
funcs(11) = "print"
funcs(12) = "end"
funcs(13) = "sqr"

If InStr(1, Command, "debug") = 0 Then
    'do not debug... (load straightaway)
Form1.Hide
Command1_Click
Command2_Click
End If

End Sub

Private Function isdelim(strin) As Boolean
isdelim = False
If strin = " " Then isdelim = True
If strin = "," Then isdelim = True
If strin = "(" Then isdelim = True
If strin = ")" Then isdelim = True
If strin = "=" Then isdelim = True
If strin = "+" Then isdelim = True
If strin = "-" Then isdelim = True
If strin = "/" Then isdelim = True
If strin = "*" Then isdelim = True

End Function

Private Sub addToken(tok As Token)
'tokens(currenttoken) = New Token
tokens(currenttoken) = tok
'Debug.Print "adding token " + tok.text
currenttoken = currenttoken + 1
ReDim Preserve tokens(currenttoken)
End Sub
Private Sub addMemory(mem As memoryentry)
Debug.Print "Adding var " & mem.text & " as " & mem.info
'tokens(currenttoken) = New Token
memory(memorytop) = mem
memorytop = memorytop + 1
ReDim Preserve memory(memorytop)
End Sub
Private Sub forstack_push(fs As forstackentry)
forstack(forstackpos) = fs
forstackpos = forstackpos + 1
ReDim Preserve forstack(forstackpos)
End Sub
Private Function forstack_pop() As forstackentry
If forstackpos = 0 Then
    MsgBox "Trying to pop empty forstack!"
    Exit Function
End If
forstackpos = forstackpos - 1
forstack_pop = forstack(forstackpos)
ReDim Preserve forstack(forstackpos)
End Function
Private Function getMemory(nam)
For a = 0 To memorytop
    If memory(a).text = nam Then
        getMemory = memory(a).info
        Exit Function
    End If
Next a
'need to allocate some
Dim m As memoryentry
m.text = nam
m.info = 0
addMemory m
getMemory = m.info
End Function
Private Function setMemory(nam, setto)
For a = 0 To memorytop
    If memory(a).text = nam Then
        memory(a).info = setto
        Exit Function
    End If
Next a
'need to allocate some
Dim m As memoryentry
m.text = nam
m.info = setto
addMemory m
End Function

Private Function getTokenType(strin) As Token
Dim tok As Token

strin = LCase(strin)
tok.text = strin
If strin = "=" Then
    tok.assignment = True
    GoTo toktypeex
End If

If strin = "+" Then
    tok.add = True
    tok.mathop = True
    GoTo toktypeex
End If

If strin = "-" Then
    tok.subtract = True
    tok.mathop = True
    GoTo toktypeex
End If

If strin = "/" Then
    tok.divide = True
    tok.mathop = True
    GoTo toktypeex
End If

If strin = "*" Then
    tok.multiply = True
    tok.mathop = True
    GoTo toktypeex
End If


If strin = "(" Then
    tok.openbracket = True
    GoTo toktypeex
End If

If strin = "," Then
    tok.delim = True
    GoTo toktypeex
End If

If strin = ")" Then
    tok.closebracket = True
    GoTo toktypeex
End If

If Len(strin) > 1 And Right(strin, 1) = ":" Then
    tok.gotolabel = True
    GoTo toktypeex
End If

'it be an ident or a func then

If IsNumeric(strin) Then
    tok.numlit = True
    tok.lit = True
End If

tok.ident = True

For q = 0 To UBound(funcs)
    If strin = funcs(q) Then tok.func = True
Next q

toktypeex:
getTokenType = tok

End Function

Private Function dofunc(fa As funcargs, argcount)
Dim fs As forstackentry

' msgbox( "text", window_type, "title" ) (for window types see vb help)
If fa.funcname = "msgbox" Then
    If argcount <> 3 Then
        MsgBox "Incorrect num of args to MsgBox!"
        Exit Function
    End If
    dofunc = MsgBox(fa.args(1), fa.args(2), fa.args(3))
    Exit Function
End If

' ifequal( var1, var2 )
If fa.funcname = "ifequal" Then
    If argcount <> 2 Then
        MsgBox "Incorrect num of args to ifequal!"
        Exit Function
    End If
    If fa.args(1) <> fa.args(2) Then
        'they are not equal
        Do
        i = i + 1
        If tokens(i).newline Then
            Exit Function
        End If
        Loop
    Else
        i = i + 1
    End If
    Exit Function
End If

' ifless( var1, var2 ) (if var2 is less than var1)
If fa.funcname = "ifless" Then
    If argcount <> 2 Then
        MsgBox "Incorrect num of args to ifless!"
        Exit Function
    End If
    If Not Val(fa.args(1)) > Val(fa.args(2)) Then
        'they are not equal
        Do
        i = i + 1
        If tokens(i).newline Then
            Exit Function
        End If
        Loop
    Else
        i = i + 1
    End If
    Exit Function
End If

' ifgreater( var1, var2 ) (if var2 is greater than var1)
If fa.funcname = "ifgreater" Then
    If argcount <> 2 Then
        MsgBox "Incorrect num of args to ifgreater!"
        Exit Function
    End If
    If Not Val(fa.args(1)) < Val(fa.args(2)) Then
        'they are not greater
        Do
        i = i + 1
        If tokens(i).newline Then
            Exit Function
        End If
        Loop
    Else
        i = i + 1
    End If
    Exit Function
End If

' inputbox( "prompt", "title", "default )
If fa.funcname = "inputbox" Then
    If argcount <> 3 Then
        MsgBox "Incorrect num of args to inputbox!"
        Exit Function
    End If
    dofunc = InputBox(fa.args(1), fa.args(2), fa.args(3))
    Exit Function
End If

' createwindow( "caption" ) (doesn't work)
If fa.funcname = "createwindow" Then
    If argcount <> 1 Then
        MsgBox "Incorrect num of args to createwindow!"
        Exit Function
    End If
    'code here sometime...
    Dim a As New Form2
    a.Caption = fa.args(1)
    a.Show
End If

' beep()
If fa.funcname = "beep" Then
    If argcount <> 0 Then
        MsgBox "Incorrect num of args to beep!"
        Exit Function
    End If
    Beep
End If

' goto( "labelname" )
If fa.funcname = "goto" Then
    If argcount <> 1 Then
        MsgBox "Incorrect num of args to goto!"
        Exit Function
    End If
    
    For z = 0 To currenttoken
        If tokens(z).gotolabel And Left(tokens(z).text, Len(fa.args(1))) = LCase(fa.args(1)) Then
            i = z
            Exit Function
        End If
    Next z
End If

' for( variable, from, to, step )
If fa.funcname = "for" Then
    If argcount <> 4 Then
        MsgBox "Incorrect num of args to For!"
        Exit Function
    End If
    fs.fortok = i
    fs.varref = fa.argrefs(1)
    fs.nfrom = fa.args(2)
    fs.nto = fa.args(3)
    fs.nstep = fa.args(4)
    setMemory fs.varref, fs.nfrom
    forstack_push fs
    Exit Function
End If

If fa.funcname = "next" Then
    fs = forstack_pop
    If Val(fs.nto) <= Val(getMemory(fs.varref)) Then
        'got to end of for... continue
        Exit Function
    End If
    setMemory fs.varref, getMemory(fs.varref) + fs.nstep
    i = fs.fortok
    forstack_push fs
    Exit Function
End If

If fa.funcname = "input" Then
    If argcount <> 0 Then
        MsgBox "Incorrect num of args to Input!"
        Exit Function
    End If
    Form2.Show
    Form2.strin = ""
    Form2.List1.AddItem "> "
    Form2.ready = False
    Do
    DoEvents
    Loop Until Form2.ready
    dofunc = Form2.strin
    Exit Function
End If

If fa.funcname = "print" Then
    If argcount <> 1 Then
        MsgBox "Incorrect num of args to Print!"
        Exit Function
    End If
'    Form2.Text1 = Form2.Text1 & fa.args(1) & vbCr
    Form2.List1.AddItem fa.args(1)
    Exit Function
End If

If fa.funcname = "end" Then
    i = currenttoken + 20
End If

If fa.funcname = "sqr" Then
    If argcount <> 1 Then
        MsgBox "Incorrect num of args to sqr!"
        Exit Function
    End If
    dofunc = Sqr(fa.args(1))
    Exit Function
End If

End Function
