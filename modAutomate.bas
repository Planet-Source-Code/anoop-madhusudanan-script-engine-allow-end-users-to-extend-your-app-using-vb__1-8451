Attribute VB_Name = "modAutomate"
'=============================================================================================================================
' MODAUTOMATE
'
'This is the module that you should customize for
'using this engine in your Application
'
'=============================================================================================================================


Public clsToAdd As clsEg


Public Function AddObjectsToSc(Sc As ScriptControl, Lview As ListView)
'Function For adding objects to Script Control


'You should set a rule while adding objects for exclding forms
'Here, I am writing code in such a way that only the forms starting
'with usf Prefix (for UserForm) are added
'Hence, the main form and edit form are not added

On Error Resume Next

'You have to load the forms before you
'add them

LoadAllUsfs True



Dim Fm As Form
    AddObjectsToList Lview

'Adding Forms To Script Control
'=================================

    For Each Fm In Forms
        If Left(CStr(Fm.Name), 3) = "usf" Then
            'Does not add the 'usf'
        addname = Right(Fm.Name, Len(Fm.Name) - 3)
            Sc.AddObject UCase(addname), Fm, True
    
        End If

    Next Fm

'Adding Classes To Script Control
'===================================
Set clsToAdd = New clsEg

Sc.AddObject "MCLASS", clsToAdd
'Adding class separately to list
Lview.ListItems.Add , , "MCLASS", 1, 1

LoadAllUsfs False

End Function



Public Function AddObjectsToList(lvView As ListView)
'Adds all the forms and controls to script control

lvView.ListItems.Clear


Dim Fm As Form, Cn As Control
On Error Resume Next


For Each Fm In Forms

If Left(CStr(Fm.Name), 3) = "usf" Then

addname = Right(Fm.Name, Len(Fm.Name) - 3)

    lvView.ListItems.Add , , UCase(addname), 1, 1
 For Each Cn In Fm.Controls
    lvView.ListItems.Add , , UCase(addname) & "." & Cn.Name, 2, 2
 Next Cn
    
End If

Next Fm



End Function



Sub LoadAllUsfs(lState As Boolean)
'Loads/Unloads all forms
'This is because script control needs a form loaded to add it

On Error Resume Next

If lState = True Then
'Load all userForms here
Load usfFORM
Else

'UnLoad all userForms here
Unload usfFORM
End If


End Sub


