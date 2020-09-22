VERSION 5.00
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "MSSCRIPT.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   Caption         =   "Script Engine"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7155
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   7155
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvView 
      Height          =   2610
      Left            =   765
      TabIndex        =   3
      Top             =   1080
      Visible         =   0   'False
      Width           =   2400
      _ExtentX        =   4233
      _ExtentY        =   4604
      View            =   1
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imlObj"
      SmallIcons      =   "imlObj"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSScriptControlCtl.ScriptControl scMain 
      Left            =   2955
      Top             =   3150
      _ExtentX        =   1005
      _ExtentY        =   1005
      AllowUI         =   -1  'True
   End
   Begin MSComctlLib.ImageList imlMain 
      Left            =   3585
      Top             =   3135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0442
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   635
      ButtonWidth     =   2408
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "imlTool"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Script"
            Key             =   "New"
            Object.ToolTipText     =   "New"
            ImageKey        =   "New"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Edit Script"
            Key             =   "Edit"
            Object.ToolTipText     =   "Edit"
            ImageKey        =   "Edit"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Delete Script"
            Key             =   "Delete"
            Object.ToolTipText     =   "Delete"
            ImageKey        =   "Delete"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Play Script"
            Key             =   "Play"
            Object.ToolTipText     =   "Play"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit Scripts"
            Key             =   "Exit Scripts"
            Object.ToolTipText     =   "Exit Application"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "About Engine"
            Key             =   "Help"
            ImageKey        =   "Help"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   4395
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   609
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlTool 
      Left            =   2310
      Top             =   3150
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0556
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0668
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":077A
            Key             =   "Forward"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":088C
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":099E
            Key             =   "Edit"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0AFA
            Key             =   "Close"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvMain 
      Height          =   2895
      Left            =   30
      TabIndex        =   2
      Top             =   405
      Width           =   7080
      _ExtentX        =   12488
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "imlMain"
      SmallIcons      =   "imlMain"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imlObj 
      Left            =   885
      Top             =   4065
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":0C56
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmScript.frx":10AA
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'=============================================================================================================================
' READ THIS BEFORE USING THE CODE:
'
' Developed by Anoop. M
' anoopj12 @ yahoo.com
'
' Anoop M, Govindanikethan, Nedumkunnam P.O, Kottayam,
' Kerala, India - 686 542
'=============================================================================================================================
'
'SCRIPTING ENGINE:
'=================
' READ THE README.TXT FILE (ATTACHED WITH THE PROJECT) BEFORE
' READING THE REST, IF YOU DON'T WANT TO GET MAD
'
'ABOUT CODE:
'=================
' You can study and view the source code for creating your
' own apps, but do not reproduce/release Icon Hunter fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
'ABOUT ME:
'=================
'
' I am a freelance programmer, and is now working as the
' Technical Advisor of an Indian Software company. I am specialising
' in Internet/Web technologies and ASP development, and is planning
' to relocate to US shortly
'
' I recently developed a broadcasting audio technology and is
' now looking for tie-ups with established US companies;
' interested in implementing their own Internet Radio Networks,
' Web Phone services and Tele-Conferencing/Voice-Chat services
' using the above technology I developed
' VISIT MY WEBSITE : http://www.geocities.com/streamingaudio for
' details regarding this technoloy
'
' Thanks for using my code
'=============================================================================================================================

'=============================================================================================================================
' FRMMAIN - THE MAIN FORM
'=============================================================================================================================

Private Sub Form_Resize()
On Error Resume Next
lvMain.Move 0, lvMain.Top, Me.ScaleWidth, Me.ScaleHeight - lvMain.Top - sbMain.Height
lvMain.ColumnHeaders(2).Width = lvMain.Width - lvMain.ColumnHeaders(1).Width - 60

End Sub

Private Sub lvMain_DblClick()
tbMain_ButtonClick tbMain.Buttons("Edit")

End Sub


Private Sub scMain_Error()

MsgBox scMain.Error.Description
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComCtlLib.Button)
Dim ItId As Long

If lvMain.ListItems.Count > 0 Then
    ItId = CLng(Right(lvMain.SelectedItem.Key, Len(lvMain.SelectedItem.Key) - 1))
End If

    On Error Resume Next
    Select Case Button.Key
        Case "Exit Scripts"
            Unload Me
            
        Case "New"
            frmEdit.DoAction "new", ItId
            mFillScripts
        Case "Edit"
            frmEdit.DoAction "edit", ItId
            mFillScripts
        Case "Delete"
            frmEdit.DoAction "delete", ItId
            mFillScripts
            
        Case "Play"
            
            'ret = MsgBox("Are you sure that you want to execute this script?", vbQuestion + vbYesNo, "Run")
            'If ret = vbNo Then Exit Sub
            
            getcode = frmEdit.DoAction("return", ItId)
            
            On Error Resume Next
            scMain.ExecuteStatement getcode
            
        Case "Help"
            
            
    End Select
    
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyReturn Then
    txtMain.Text = UCase(txtMain.Text)
    txtMain.SelStart = Len(txtMain.Text)
End If

End Sub

Private Sub Form_Load()

mFillScripts


'Adding the script control itself to the script
scMain.AddObject "SCRIPT", scMain, True

AddObjectsToSc scMain, lvView

End Sub

Sub mFillScripts()
On Error Resume Next

If FillListWithRecords(gDatabasename, "Select * from scripts", lvMain, 1, 1) Then
On Error Resume Next
    lvMain.ColumnHeaders(1).Width = lvMain.ColumnHeaders(1).Width * 2
    lvMain.ColumnHeaders(2).Width = lvMain.Width - lvMain.ColumnHeaders(1).Width - 60
Else
    MsgBox "Unable to open scripts database. Please exit and try again", vbCritical, "Error"
End If

If lvMain.ListItems.Count > 0 Then
     lvMain.ListItems(1).Selected = True
    For i = 1 To tbMain.Buttons.Count
        tbMain.Buttons(i).Enabled = True
    Next i
Else
    tbMain.Buttons("Edit").Enabled = False
    tbMain.Buttons("Delete").Enabled = False
    tbMain.Buttons("Play").Enabled = False
        
End If

sbMain.SimpleText = lvMain.ListItems.Count & " Scripts Found. Double click to exectute a script"
End Sub




