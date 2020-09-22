Attribute VB_Name = "modMain"
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
'MODMAIN
'
'Main Module
'=============================================================================================================================

Public gDatabasename As String

Sub Main()
'Our saga starts from here

    gDatabasename = App.Path & "\Scriptdb.mdb"
    frmMain.Show
End Sub
