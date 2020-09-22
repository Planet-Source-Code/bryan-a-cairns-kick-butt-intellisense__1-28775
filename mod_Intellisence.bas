Attribute VB_Name = "mod_Intellisence"
'// THESE ARE USED FOR OBTAIANING EXACT LOCATION OF     //
'// CURSOR POSITION IN RTF BOX                          //

    Public Declare Function GetCaretPos Lib "user32" _
                   (lpPoint As POINTAPI) As Long

    Type POINTAPI
        X As Long
        Y As Long
    End Type

    Public POINTAPI As POINTAPI
    
'// THIS PUBLIC VARIABLE HOLDS THE "OBJECT" -- THAT IS, //
'// THE STRING BEFORE THE "DOT". EXAMPLE. IF YOU        //
'// TYPE ME.CAPTION, "ME" IS WHAT gObjName HOLDS        //

    Public gObjName As String
    

'// THIS API CALL IS USED TO ALLOW YOU TO VIEW THE      //
'// TEXT FILE IN NOTEPAD. THIS CALL IS NOT NEEDED       //
'// FOR THE INTELLISENSE FUNCTIONALITY                  //

    Declare Function ShellExecute _
            Lib "shell32.dll" Alias "ShellExecuteA" _
            (ByVal hwnd As Long, _
             ByVal lpOperation As String, _
             ByVal lpFile As String, _
             ByVal lpParameters As String, _
             ByVal lpDirectory As String, _
             ByVal nShowCmd As Long) As Long
             



