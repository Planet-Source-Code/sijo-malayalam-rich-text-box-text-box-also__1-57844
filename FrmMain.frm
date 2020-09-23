VERSION 5.00
Object = "{C90FA0A5-D6F9-4F15-9364-F77385D88F48}#13.0#0"; "SSMLMOCX.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "SIJO Soft Malayalam Demonstration"
   ClientHeight    =   8355
   ClientLeft      =   150
   ClientTop       =   855
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dlg 
      Left            =   7395
      Top             =   3645
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SSMLMOCX.CtlSSMLM mlm 
      Height          =   8340
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   14711
   End
   Begin VB.Menu MnuFile 
      Caption         =   "&File"
      Begin VB.Menu MnuFileOpenTF 
         Caption         =   "Open Text File"
      End
      Begin VB.Menu MnuFileOpenTRF 
         Caption         =   "Open Rtf File"
      End
      Begin VB.Menu MnuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileSaveAsTF 
         Caption         =   "Save As Text File"
      End
      Begin VB.Menu MnuFileSaveAsRF 
         Caption         =   "Save As Rtf Text File"
      End
      Begin VB.Menu MnuFileSaveAsWP 
         Caption         =   "Save As WebPage"
      End
      Begin VB.Menu MnuBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu MnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu MnuEditCut 
         Caption         =   "Cut"
         Shortcut        =   ^X
      End
      Begin VB.Menu MnuEditCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu MnuEditPaste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu MnuInsert 
      Caption         =   "&Insert"
      Begin VB.Menu MnuInsertDate 
         Caption         =   "Insert Date"
      End
      Begin VB.Menu MnuInsertName 
         Caption         =   "Insert Name"
      End
   End
   Begin VB.Menu MnuFormat 
      Caption         =   "Fo&rmat"
      Begin VB.Menu MnuFormatFont 
         Caption         =   "Font"
      End
      Begin VB.Menu MnuFormatBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFormatFB 
         Caption         =   "Set Font Bold"
      End
      Begin VB.Menu MnuFormatFI 
         Caption         =   "Set Font Italic"
      End
      Begin VB.Menu MnuFormatFU 
         Caption         =   "Set Fornt Underline"
      End
      Begin VB.Menu MnuFormatBar1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFormatFontColor 
         Caption         =   "Font Color"
      End
      Begin VB.Menu MnuFormatBar2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuFormatWrodWrap 
         Caption         =   "WordWrap"
      End
      Begin VB.Menu MnuFormatBar3 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuSwitch 
      Caption         =   "Switch"
      Begin VB.Menu MnuSwitchShowMLMKB 
         Caption         =   "Show Malayalam Keyboard"
      End
      Begin VB.Menu MnuSwitchHideMLMKB 
         Caption         =   "Hide Malayalam Keyboard"
      End
      Begin VB.Menu MnuSwitchBar0 
         Caption         =   "-"
      End
      Begin VB.Menu MnuSwitchMLMLg 
         Caption         =   "Switch to Malayalam Language"
      End
      Begin VB.Menu MnuSwitchEngLg 
         Caption         =   "Switch to English Language"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu MnuOffer 
         Caption         =   "Offer"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    App.TaskVisible = False
    mlm.About
End Sub

Private Sub Form_Resize()
    mlm.Width = Me.Width - 150
    mlm.Height = Me.Height - 700
End Sub

Private Sub MnuAbout_Click()
mlm.About
End Sub

Private Sub MnuEditCopy_Click()
    mlm.Copy
End Sub

Private Sub MnuEditCut_Click()
    mlm.Cut
End Sub

Private Sub MnuEditPaste_Click()
    mlm.Paste
End Sub

Private Sub MnuFileOpenTF_Click()
With Dlg
    .DialogTitle = "Open RTF File"
    .Filter = "Text File(*.txt|*.txt"
    .ShowOpen
    If Not .FileName = "" Then
        mlm.OpenText .FileName
    End If
End With
End Sub

Private Sub MnuFileOpenTRF_Click()
With Dlg
    .DialogTitle = "Open RTF File"
    .Filter = "Rich Text File(*.rtf|*.rtf"
    .ShowOpen
    If Not .FileName = "" Then
        mlm.OpenRtf .FileName
    End If
End With
End Sub

Private Sub MnuFileOpenWord_Click()
'With Dlg
    '.DialogTitle = "Open Word File"
    '.Filter = "Word File(*.doc|*.doc"
    'Dlg.ShowSave
    'If Not Dlg.FileName = "" Then
        'mlm.SaveAsWordFile Dlg.FileName
    'End If
'End With
End Sub

Private Sub MnuFileSaveAsRF_Click()
With Dlg
    .Filter = "Rich Text File(*.rtf|*.rtf"
    Dlg.ShowSave
    If Not Dlg.FileName = "" Then
        mlm.SaveAsRtf Dlg.FileName
    End If
End With
End Sub

Private Sub MnuFileSaveAsTF_Click()
With Dlg
    .Filter = "Text File(*.txt|*.txt"
    Dlg.ShowSave
    If Not Dlg.FileName = "" Then
        mlm.SaveAsText Dlg.FileName
    End If
End With
End Sub

Private Sub MnuFileSaveAsWF_Click()
'With Dlg
    '.Filter = "Word File(*.doc|*.doc"
    'Dlg.ShowSave
    'If Not Dlg.FileName = "" Then
        'mlm.SaveAsWordFile Dlg.FileName
    'End If
'End With
End Sub

Private Sub MnuFileSaveAsWP_Click()
With Dlg
    .Filter = "Web Page(*.htm|*.htm"
    Dlg.ShowSave
    If Not Dlg.FileName = "" Then
        mlm.SaveAsWebPage Dlg.FileName
    End If
End With
End Sub

Private Sub MnuFormatFB_Click()
    mlm.SetFontBold (1)
End Sub

Private Sub MnuFormatFI_Click()
    mlm.SetFontItalic (1)
End Sub

Private Sub MnuFormatFont_Click()
    With Dlg '- Common Dialog Control
        .ShowFont
        mlm.MalayalamFontName = Dlg.FontName
        '(Must be a SIJO Soft Malyalam Enabled Malyalam Font)
    End With
End Sub

Private Sub MnuFormatFontColor_Click()
    With Dlg '- Common Dialog Control
        Dlg.ShowColor
        mlm.FontColor = Dlg.Color
    End With
End Sub

Private Sub MnuFormatFU_Click()
    mlm.SetFontUnderline (1)
End Sub

Private Sub MnuFormatWrodWrap_Click()
    mlm.WordWrap = 1 '- True
   'mlm.WordWrap = 0 '- false
End Sub

Private Sub MnuInsertDate_Click()
    mlm.Text = mlm.Text & " " & Date
End Sub

Private Sub MnuInsertName_Click()
    mlm.Text = mlm.Text & " " & "Øßæ¼Þ çØÞËí¿í ÎÜÏÞ{¢"
End Sub

Private Sub MnuOffer_Click()
mlm.Text = "ÎÜÏÞ{ßµZAí §ì µYçd¿ÞZ §çMÞZ Øì¼ÈcÎÞÏß" & vbCrLf & _
"ÜÍßAáKÄÞÃíµâ¿áÄW ÕßÕø¹ZAí" & vbCrLf & "Øßç¼Þ.®Gí æµ.çµÞ¢ ØwVÖßAáµ" & vbCrLf
End Sub

Private Sub MnuSwitchEngLg_Click()
    mlm.DefaultLanguage = 0 '- English Language
End Sub

Private Sub MnuSwitchHideMLMKB_Click()
    mlm.HideMalayalamKeyboard
End Sub

Private Sub MnuSwitchMLMLg_Click()
    mlm.DefaultLanguage = 1 ' - Malayalam Language
End Sub

Private Sub MnuSwitchShowMLMKB_Click()
    mlm.ShowMalayalamKeyboard
End Sub
