VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNetSend 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Net Send Cheater"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8850
   Icon            =   "frmNetSend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   8850
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      Height          =   1665
      Index           =   5
      Left            =   0
      ScaleHeight     =   1605
      ScaleWidth      =   8595
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7080
      Width           =   8655
      Begin RichTextLib.RichTextBox rtb 
         Height          =   1650
         Index           =   0
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   2910
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmNetSend.frx":08CA
      End
   End
   Begin RichTextLib.RichTextBox rtbINI 
      Height          =   495
      Left            =   0
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   873
      _Version        =   393217
      TextRTF         =   $"frmNetSend.frx":094C
   End
   Begin VB.PictureBox pic 
      Height          =   1665
      Index           =   4
      Left            =   0
      ScaleHeight     =   1605
      ScaleWidth      =   8595
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5400
      Width           =   8655
      Begin VB.CommandButton cmd 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   0
         TabIndex        =   14
         Top             =   1250
         Width           =   8535
      End
      Begin VB.ListBox li 
         Height          =   1230
         Index           =   0
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   13
         Top             =   0
         Width           =   8535
      End
   End
   Begin VB.PictureBox pic 
      Height          =   1665
      Index           =   3
      Left            =   0
      ScaleHeight     =   1605
      ScaleWidth      =   8595
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3720
      Width           =   8655
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   2
         Left            =   5040
         Sorted          =   -1  'True
         TabIndex        =   28
         Top             =   0
         Width           =   3495
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Index           =   6
         Left            =   4320
         TabIndex        =   31
         Top             =   1250
         Width           =   4215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Index           =   5
         Left            =   0
         TabIndex        =   30
         Top             =   1250
         Width           =   4215
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   26
         Top             =   0
         Width           =   3495
      End
      Begin MSComctlLib.ListView lv 
         Height          =   855
         Index           =   1
         Left            =   0
         TabIndex        =   29
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label lbl 
         Caption         =   "Name:"
         Height          =   255
         Index           =   5
         Left            =   4320
         TabIndex        =   27
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Group:"
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   25
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox pic 
      Height          =   1665
      Index           =   2
      Left            =   0
      ScaleHeight     =   1605
      ScaleWidth      =   8595
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2040
      Width           =   8655
      Begin VB.CommandButton cmd 
         Caption         =   "&Delete"
         Enabled         =   0   'False
         Height          =   375
         Index           =   4
         Left            =   4320
         TabIndex        =   22
         Top             =   1250
         Width           =   4215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Add"
         Enabled         =   0   'False
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   21
         Top             =   1250
         Width           =   4215
      End
      Begin MSComctlLib.ListView lv 
         Height          =   855
         Index           =   0
         Left            =   0
         TabIndex        =   20
         Top             =   360
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   1508
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   1
         Left            =   5040
         TabIndex        =   19
         Top             =   0
         Width           =   3495
      End
      Begin VB.TextBox txt 
         Height          =   315
         Index           =   0
         Left            =   720
         TabIndex        =   17
         Top             =   0
         Width           =   3495
      End
      Begin VB.Label lbl 
         Caption         =   "Computer:"
         Height          =   255
         Index           =   3
         Left            =   4280
         TabIndex        =   18
         Top             =   120
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "Name:"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   735
      End
   End
   Begin VB.PictureBox pic 
      Height          =   1665
      Index           =   1
      Left            =   0
      ScaleHeight     =   1605
      ScaleWidth      =   8595
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   360
      Width           =   8655
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "Sh&ow Sent Messages:"
         Height          =   255
         Index           =   1
         Left            =   6560
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   840
         Width           =   1935
      End
      Begin VB.CheckBox chk 
         Alignment       =   1  'Right Justify
         Caption         =   "S&ave to Quick Messages:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   800
         Width           =   2175
      End
      Begin VB.CommandButton cmd 
         Cancel          =   -1  'True
         Caption         =   "&Close"
         Height          =   375
         Index           =   1
         Left            =   4320
         TabIndex        =   9
         Top             =   1250
         Width           =   4215
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Send"
         Default         =   -1  'True
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   8
         Top             =   1250
         Width           =   4215
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   1
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   360
         Width           =   7815
      End
      Begin VB.ComboBox cmb 
         Height          =   315
         Index           =   0
         Left            =   720
         Sorted          =   -1  'True
         TabIndex        =   3
         Top             =   0
         Width           =   7815
      End
      Begin VB.Label lbl 
         Caption         =   "Message:"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   4
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lbl 
         Caption         =   "To:"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   120
         Width           =   735
      End
   End
   Begin MSComctlLib.TabStrip ts 
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   661
      MultiRow        =   -1  'True
      TabMinWidth     =   2117
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Send &Message"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Persons"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Groups"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Quick Messages"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&Help"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmNetSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sINI As String

Private Sub ReadINI()
    Dim sTemp1 As String
    Dim sTemp2 As String
    Dim sTemp3 As String
    Dim bTF As Boolean
    
    Open sINI For Input As #1
    For Counter = 0 To li.Count - 1
        li(Counter).Clear
        lv(Counter).ListItems.Clear
        lv(Counter).ColumnHeaders.Clear
        lv(Counter).Sorted = True
        lv(Counter).Checkboxes = False
        lv(Counter).View = lvwReport
    Next
    lv(0).ColumnHeaders.Add , , "Name", 4110
    lv(0).ColumnHeaders.Add , , "Computer", 4110
    lv(1).ColumnHeaders.Add , , "Group", 4110
    lv(1).ColumnHeaders.Add , , "Name", 4110
    Do While Not EOF(1)
        Input #1, sTemp1, sTemp2, sTemp3
        If sTemp1 = "Top" Then
            Me.Top = Val(sTemp2)
        ElseIf sTemp1 = "Left" Then
            Me.Left = Val(sTemp2)
        ElseIf sTemp1 = "ShowSent" Then
            chk(1).Value = Val(sTemp2)
        ElseIf sTemp1 = "Message" Then
            li(0).AddItem sTemp2
        ElseIf sTemp1 = "Person" Then
            Set itmX = lv(0).ListItems.Add(, , sTemp2)
            itmX.SubItems(1) = sTemp3
        ElseIf sTemp1 = "Group" Then
            Set itmX = lv(1).ListItems.Add(, , sTemp2)
            itmX.SubItems(1) = sTemp3
        End If
    Loop
    Close #1
    cmb(0).Clear
    cmb(2).Clear
    For Counter = 1 To lv(0).ListItems.Count
        cmb(0).AddItem lv(0).ListItems(Counter).Text + "-P"
        cmb(2).AddItem lv(0).ListItems(Counter).Text
    Next
    For Counter = 1 To lv(1).ListItems.Count
        bTF = False
        For Counter1 = 0 To cmb(0).ListCount - 1
            If UCase(cmb(0).List(Counter1)) = UCase(lv(1).ListItems(Counter).Text + "-G") Then bTF = True
        Next
        If bTF = False Then cmb(0).AddItem lv(1).ListItems(Counter).Text + "-G"
    Next
    cmb(1).Clear
    For Counter = 0 To li(0).ListCount - 1
        cmb(1).AddItem li(0).List(Counter)
    Next
End Sub

Private Sub WriteINI()
    rtbINI.Text = ""
'top
    rtbINI.Text = rtbINI.Text + "Top," + LTrim(RTrim(Str(Me.Top))) + "," + vbCrLf
'left
    rtbINI.Text = rtbINI.Text + "Left," + LTrim(RTrim(Str(Me.Left))) + "," + vbCrLf
'ShowSent
    rtbINI.Text = rtbINI.Text + "ShowSent," + LTrim(RTrim(Str(chk(1).Value))) + "," + vbCrLf
'messages
    For Counter = 0 To li(0).ListCount - 1
        rtbINI.Text = rtbINI.Text + "Message," + li(0).List(Counter) + "," + vbCrLf
    Next
'person
    For Counter = 1 To lv(0).ListItems.Count
        rtbINI.Text = rtbINI.Text + "Person," + lv(0).ListItems(Counter).Text + "," + lv(0).ListItems(Counter).SubItems(1) + vbCrLf
    Next
'person
    For Counter = 1 To lv(1).ListItems.Count
        rtbINI.Text = rtbINI.Text + "Group," + lv(1).ListItems(Counter).Text + "," + lv(1).ListItems(Counter).SubItems(1) + vbCrLf
    Next

'save INI
    rtbINI.SaveFile sINI, 1
End Sub

Private Sub cmb_Change(Index As Integer)
    If Index = 0 Then
        If Len(cmb(0).Text) > 0 And Len(cmb(1).Text) > 0 Then
            cmd(0).Enabled = True
        Else
            cmd(0).Enabled = False
        End If
    ElseIf Index = 1 Then
        If Len(cmb(0).Text) > 0 And Len(cmb(1).Text) > 0 Then
            cmd(0).Enabled = True
        Else
            cmd(0).Enabled = False
        End If
    ElseIf Index = 2 Then
        If Len(txt(2).Text) > 0 And Len(cmb(2).Text) > 0 Then
            cmd(5).Enabled = True
        Else
            cmd(5).Enabled = False
        End If
    End If
End Sub

Private Sub cmb_Click(Index As Integer)
    cmb_Change Index
End Sub

Private Sub cmd_Click(Index As Integer)
    Dim sCommand As String
    Dim RetVal
    Dim bTF As Boolean
    Dim lYN As Long
    Dim sYN As String
    Dim sTemp As String
'    On Error Resume Next

    If Index = 0 Then       'send send
        If Len(cmb(0).Text) < 1 Then
            MsgBox "You must enter a Machine Name in the To field.", vbExclamation
            Exit Sub
        Else
            If Len(cmb(1).Text) < 1 Then
                MsgBox "You must enter a Message in the Message field.", vbExclamation
                Exit Sub
            Else
                If UCase(Right(cmb(0).Text, 2)) = "-P" Then
                    For Counter = 1 To lv(0).ListItems.Count
                        If UCase(cmb(0).Text) = UCase(lv(0).ListItems(Counter).Text + "-P") Then
                            sTemp = lv(0).ListItems(Counter).SubItems(1)
                        End If
                    Next
                    sCommand = "NET SEND " + sTemp + " " + Chr(34) + cmb(1).Text + Chr(34)
                    If chk(1).Value = 1 Then MsgBox sCommand
                    RetVal = Shell(sCommand, vbHide)
                ElseIf UCase(Right(cmb(0).Text, 2)) = "-G" Then
                    For Counter = 1 To lv(1).ListItems.Count
                        If UCase(cmb(0).Text) = UCase(lv(1).ListItems(Counter).Text + "-G") Then
                            sTemp = ""
                            sTemp = lv(1).ListItems(Counter).SubItems(1)
                            For Counter1 = 1 To lv(0).ListItems.Count
                                If UCase(sTemp) = UCase(lv(0).ListItems(Counter1).Text) Then
                                    sTemp = lv(0).ListItems(Counter1).SubItems(1)
                                    Exit For
                                End If
                            Next
                            If Len(sTemp) > 0 Then
                                sCommand = "NET SEND " + sTemp + " " + Chr(34) + cmb(1).Text + Chr(34)
                                If chk(1).Value = 1 Then MsgBox sCommand
                                RetVal = Shell(sCommand, vbHide)
                            End If
                        End If
                    Next
                Else
                    sCommand = "NET SEND " + cmb(0).Text + " " + Chr(34) + cmb(1).Text + Chr(34)
                    If chk(1).Value = 1 Then MsgBox sCommand
                    RetVal = Shell(sCommand, vbHide)
                End If
                If chk(0).Value = 1 Then
                    bTF = False
                    For Counter = 0 To li(0).ListCount - 1
                        If UCase(cmb(1).Text) = UCase(li(0).List(Counter)) Then bTF = True
                    Next
                    If bTF = False Then li(0).AddItem cmb(1).Text
                    cmb(1).Clear
                    For Counter = 0 To li(0).ListCount - 1
                        cmb(1).AddItem li(0).List(Counter)
                    Next
                Else
                    cmb(1).Text = ""
                End If
            End If
        End If
    ElseIf Index = 1 Then   'send close
        Unload Me
    ElseIf Index = 2 Then   'message delete
        sYN = "Delete " + li(0).Text + "?"
        lYN = MsgBox(sYN, vbYesNo + vbDefaultButton2)
        If lYN = vbYes Then
            li(0).RemoveItem li(0).ListIndex
            cmd(2).Enabled = False
        End If
        cmb(1).Clear
        For Counter = 0 To li(0).ListCount - 1
            cmb(1).AddItem li(0).List(Counter)
        Next
    ElseIf Index = 3 Then   'persons add
        If UCase(Right(txt(0).Text, 2)) = "-P" Or UCase(Right(txt(0).Text, 2)) = "-G" Then
            MsgBox "Name can not end in '-P' or '-G'.", vbExclamation
            txt(0).Text = ""
            txt(0).SetFocus
            Exit Sub
        End If
        For Counter = 1 To lv(0).ListItems.Count
            If UCase(txt(0).Text) = UCase(lv(0).ListItems(Counter).Text) Then
                MsgBox "The name " + txt(0).Text + " already exists.", vbExclamation
                txt(0).Text = ""
                txt(0).SetFocus
                Exit Sub
            End If
        Next
        Set itmX = lv(0).ListItems.Add(, , txt(0).Text)
        itmX.SubItems(1) = txt(1).Text
        txt(0).Text = ""
        txt(1).Text = ""
        txt(0).SetFocus
        cmb(0).Clear
        cmb(2).Clear
        For Counter = 1 To lv(0).ListItems.Count
            cmb(0).AddItem lv(0).ListItems(Counter).Text + "-P"
            cmb(2).AddItem lv(0).ListItems(Counter).Text
        Next
        For Counter = 1 To lv(1).ListItems.Count
            bTF = False
            For Counter1 = 0 To cmb(0).ListCount - 1
                If UCase(cmb(0).List(Counter1)) = UCase(lv(1).ListItems(Counter).Text + "-G") Then bTF = True
            Next
            If bTF = False Then cmb(0).AddItem lv(1).ListItems(Counter).Text + "-G"
        Next
        cmd(4).Enabled = True
    ElseIf Index = 4 Then   'persons delete
        sYN = "Delete " + lv(0).SelectedItem.Text + "?"
        lYN = MsgBox(sYN, vbYesNo + vbDefaultButton2)
        If lYN = vbYes Then
            sTemp = lv(0).SelectedItem.Text
            lv(0).ListItems.Remove lv(0).SelectedItem.Index
again:
            bTF = False
            For Counter = 1 To lv(1).ListItems.Count
                If UCase(lv(1).ListItems(Counter).SubItems(1)) = UCase(sTemp) Then
                    lv(1).ListItems.Remove Counter
                    bTF = True
                    Exit For
                End If
            Next
        End If
        If bTF = True Then GoTo again
        txt(0).Text = ""
        txt(1).Text = ""
        txt(0).SetFocus
        cmb(0).Clear
        cmb(2).Clear
        For Counter = 1 To lv(0).ListItems.Count
            cmb(0).AddItem lv(0).ListItems(Counter).Text + "-P"
            cmb(2).AddItem lv(0).ListItems(Counter).Text
        Next
        For Counter = 1 To lv(1).ListItems.Count
            bTF = False
            For Counter1 = 0 To cmb(0).ListCount - 1
                If UCase(cmb(0).List(Counter1)) = UCase(lv(1).ListItems(Counter).Text + "-G") Then bTF = True
            Next
            If bTF = False Then cmb(0).AddItem lv(1).ListItems(Counter).Text + "-G"
        Next
        If lv(0).ListItems.Count < 1 Then
            cmd(4).Enabled = False
            txt(0).SetFocus
            Exit Sub
        End If
    ElseIf Index = 5 Then   'group add
        bTF = False
        For Counter = 0 To cmb(2).ListCount - 1
            If UCase(cmb(2).Text) = UCase(cmb(2).List(Counter)) Then
                bTF = True
                Exit For
            End If
        Next
        If bTF = False Then
            MsgBox "Not a valid Person Name.  Select one from the list.", vbExclamation
            cmb(2).Text = ""
            cmb(2).SetFocus
            Exit Sub
        End If
        If UCase(Right(txt(2).Text, 2)) = "-P" Or UCase(Right(txt(2).Text, 2)) = "-G" Then
            MsgBox "Group can not end in '-P' or '-G'.", vbExclamation
            txt(2).Text = ""
            txt(2).SetFocus
            Exit Sub
        End If
        Set itmX = lv(1).ListItems.Add(, , txt(2).Text)
        itmX.SubItems(1) = cmb(2).Text
        cmb(2).Text = ""
        txt(2).SetFocus
        cmb(0).Clear
        cmb(2).Clear
        For Counter = 1 To lv(0).ListItems.Count
            cmb(0).AddItem lv(0).ListItems(Counter).Text + "-P"
            cmb(2).AddItem lv(0).ListItems(Counter).Text
        Next
        For Counter = 1 To lv(1).ListItems.Count
            bTF = False
            For Counter1 = 0 To cmb(0).ListCount - 1
                If UCase(cmb(0).List(Counter1)) = UCase(lv(1).ListItems(Counter).Text + "-G") Then bTF = True
            Next
            If bTF = False Then cmb(0).AddItem lv(1).ListItems(Counter).Text + "-G"
        Next
        cmd(4).Enabled = True
    ElseIf Index = 6 Then   'group delete
        sYN = "Delete " + lv(1).SelectedItem.Text + "?"
        lYN = MsgBox(sYN, vbYesNo + vbDefaultButton2)
        If lYN = vbYes Then
            lv(1).ListItems.Remove lv(1).SelectedItem.Index
        End If
        txt(2).Text = ""
        cmb(2).Text = ""
        txt(2).SetFocus
        cmb(0).Clear
        cmb(2).Clear
        For Counter = 1 To lv(0).ListItems.Count
            cmb(0).AddItem lv(0).ListItems(Counter).Text + "-P"
            cmb(2).AddItem lv(0).ListItems(Counter).Text
        Next
        For Counter = 1 To lv(1).ListItems.Count
            bTF = False
            For Counter1 = 0 To cmb(0).ListCount - 1
                If UCase(cmb(0).List(Counter1)) = UCase(lv(1).ListItems(Counter).Text + "-G") Then bTF = True
            Next
            If bTF = False Then cmb(0).AddItem lv(1).ListItems(Counter).Text + "-G"
        Next
        If lv(1).ListItems.Count < 1 Then
            cmd(6).Enabled = False
            txt(2).SetFocus
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim sPath As String
    Dim sHelp As String
    Dim bs
    Set bs = CreateObject("scripting.filesystemobject")
    sPath = App.Path
'Visuals
    Me.Width = 8940
    Me.Height = 2475
    ts.Height = 2055
    For Counter = 1 To 5
        pic(Counter).Left = 100
        pic(Counter).Top = 360
        pic(Counter).Height = 1665
        pic(Counter).Width = 8645
        pic(Counter).BorderStyle = 0
        pic(Counter).Visible = False
    Next
    pic(1).Visible = True
    Me.Show
    cmb(0).SetFocus
'load defaults
    If Right(sPath, 1) <> "\" Then sPath = sPath + "\"
    sINI = sPath + "NetSend.ini"
    If bs.FileExists(sINI) = True Then ReadINI
    sHelp = sPath + "NetSend.rtf"
    If bs.FileExists(sHelp) = True Then
        rtb(0).LoadFile sHelp, 0
    Else
        rtb(0).Text = "Help File not found."
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    WriteINI
End Sub

Private Sub li_Click(Index As Integer)
    If Index = 0 Then
        If li(0).ListCount > 0 Then
            cmd(2).Enabled = True
        Else
            cmd(2).Enabled = False
        End If
    End If
End Sub

Private Sub lv_Click(Index As Integer)
    If Index = 0 Then       'persons
        If lv(0).ListItems.Count > 0 Then
            cmd(4).Enabled = True
        Else
            cmd(4).Enabled = False
        End If
    ElseIf Index = 1 Then   'groups
        If lv(1).ListItems.Count > 0 Then
            cmd(6).Enabled = True
        Else
            cmd(6).Enabled = False
        End If
    End If
End Sub

Private Sub ts_Click()
    For Counter = 1 To 5
        pic(Counter).Visible = False
    Next
    pic(ts.SelectedItem.Index).Visible = True
    If ts.SelectedItem.Index = 1 Then       'send message
        cmd(0).Default = True
        cmd(1).Cancel = True
    ElseIf ts.SelectedItem.Index = 2 Then   'persons
        cmd(3).Default = True
        If lv(0).ListItems.Count > 0 Then
            cmd(4).Enabled = True
        Else
            cmd(4).Enabled = False
        End If
    ElseIf ts.SelectedItem.Index = 3 Then   'groups
        cmd(5).Default = True
        If lv(1).ListItems.Count > 0 Then
            cmd(6).Enabled = True
        Else
            cmd(6).Enabled = False
        End If
    ElseIf ts.SelectedItem.Index = 4 Then   'quick messages
    ElseIf ts.SelectedItem.Index = 5 Then   'help
    End If
End Sub

Private Sub txt_Change(Index As Integer)
    If Index = 0 Then       'persons name
        If Len(txt(0).Text) > 0 And Len(txt(1).Text) > 0 Then
            cmd(3).Enabled = True
        Else
            cmd(3).Enabled = False
        End If
    ElseIf Index = 1 Then   'persons computer
        If Len(txt(0).Text) > 0 And Len(txt(1).Text) > 0 Then
            cmd(3).Enabled = True
        Else
            cmd(3).Enabled = False
        End If
    ElseIf Index = 2 Then   'group
        If Len(txt(2).Text) > 0 And Len(cmb(2).Text) > 0 Then
            cmd(5).Enabled = True
        Else
            cmd(5).Enabled = False
        End If
    End If
End Sub
