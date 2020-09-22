VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Script Editor"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8265
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   8265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fr 
      Caption         =   "Objects"
      Height          =   3015
      Index           =   2
      Left            =   5580
      TabIndex        =   9
      Top             =   1515
      Width           =   2580
      Begin MSComctlLib.ListView lvView 
         Height          =   2610
         Left            =   90
         TabIndex        =   10
         Top             =   270
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
      Begin MSComctlLib.ImageList imlObj 
         Left            =   0
         Top             =   0
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
               Picture         =   "frmEdit.frx":014A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEdit.frx":059E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Data datMain 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   480
      Left            =   1095
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Scripts"
      Top             =   4215
      Visible         =   0   'False
      Width           =   2475
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Cancel"
      Height          =   405
      Left            =   6855
      TabIndex        =   8
      Top             =   750
      Width           =   1260
   End
   Begin VB.CommandButton cmdDone 
      Caption         =   "&Done"
      Height          =   405
      Left            =   6855
      TabIndex        =   7
      Top             =   195
      Width           =   1260
   End
   Begin VB.Frame fr 
      Caption         =   "Code"
      Height          =   3030
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   1515
      Width           =   5370
      Begin VB.TextBox txtFld 
         DataField       =   "Code"
         DataSource      =   "datMain"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Index           =   2
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   270
         Width           =   5100
      End
   End
   Begin VB.Frame fr 
      Height          =   1290
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   6315
      Begin VB.TextBox txtFld 
         DataField       =   "Remark"
         DataSource      =   "datMain"
         Height          =   315
         Index           =   1
         Left            =   1365
         MaxLength       =   250
         TabIndex        =   4
         Top             =   750
         Width           =   4725
      End
      Begin VB.TextBox txtFld 
         DataField       =   "Name"
         DataSource      =   "datMain"
         Height          =   315
         Index           =   0
         Left            =   1365
         MaxLength       =   40
         TabIndex        =   2
         Top             =   300
         Width           =   4725
      End
      Begin VB.Label lblFld 
         Caption         =   "Remark         :"
         Height          =   285
         Index           =   1
         Left            =   90
         TabIndex        =   3
         Top             =   780
         Width           =   1260
      End
      Begin VB.Label lblFld 
         Caption         =   "Script Name  :"
         Height          =   285
         Index           =   0
         Left            =   90
         TabIndex        =   1
         Top             =   375
         Width           =   1260
      End
   End
End
Attribute VB_Name = "frmEdit"
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
' FRMEDIT.FRM - FOR EDITING/CREATING SCRIPTS
'=============================================================================================================================


Public Function DoAction(Action As String, Id As Long) As String
datMain.DatabaseName = gDatabasename
datMain.Refresh

Load Me

'For getting the exact count
On Error Resume Next
datMain.Recordset.MoveLast
datMain.Recordset.MoveFirst

'Do the action

    Select Case LCase(Action)
        
        Case "new"
            datMain.Recordset.AddNew
            Me.Caption = "Script Editor - New Script"
            Me.Show vbModal
        
        Case "edit"
            If LocateItem(Id) Then
                datMain.Recordset.Edit
                Me.Caption = "Script Editor - Edit Script"
                Me.Show vbModal
            Else
                Unload Me
                Exit Function
            End If
        
        Case "delete"
            If LocateItem(Id) Then
                datMain.Recordset.Delete
                Unload Me
            Else
                Exit Function
            End If
                        
        Case "return"
            If LocateItem(Id) Then
                DoAction = txtFld(2).Text
                Unload Me
            Else
                DoAction = ""
            End If
                        
    End Select

End Function


Private Sub cmdDone_Click()
On Error GoTo noupdate
datMain.Recordset.Update
Unload Me

Exit Sub

noupdate:
MsgBox "Unable to update the script. " & Err.Description, vbCritical, "Unable To Update"
End Sub

Private Sub cmdExit_Click()
On Error Resume Next
datMain.Recordset.CancelUpdate
Unload Me

End Sub

Function LocateItem(Id As Long) As Boolean

With datMain.Recordset
    For i = 0 To .RecordCount - 1
        If .Fields(0).Value = Id Then
            LocateItem = True
            Exit Function
        End If
        .MoveNext
    Next i
    
LocateItem = True
    
End With

End Function

Private Sub Form_Load()
lvView.ListItems.Clear

On Error Resume Next

With frmMain.lvView
    For i = 1 To .ListItems.Count
        lvView.ListItems.Add , , .ListItems(i).Text, .ListItems(i).Icon, .ListItems(i).SmallIcon
    Next i
End With

End Sub

Private Sub lvView_DblClick()
txtFld(2).SelText = lvView.SelectedItem.Text
End Sub

Private Sub txtFld_GotFocus(Index As Integer)
If Index <> 2 Then
    txtFld(Index).SelStart = 0
    txtFld(Index).SelLength = Len(txtFld(Index).Text)
End If
End Sub

Sub MakeCap()
'Write some routine to capitalize the objects in the textbox

End Sub
