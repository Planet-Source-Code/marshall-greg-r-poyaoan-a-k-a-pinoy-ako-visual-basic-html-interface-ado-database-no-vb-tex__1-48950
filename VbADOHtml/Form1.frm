VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form frmMain 
   Caption         =   "Simple ADO database in Web Browser."
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   6090
   ScaleWidth      =   8970
   StartUpPosition =   2  'CenterScreen
   Begin SHDocVwCtl.WebBrowser Web1 
      Height          =   1575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   2778
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.ListBox lstBox 
      Height          =   450
      Left            =   1800
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'   Please be considerate if my English is not perfect. I'm from Philippines.
'   I commented almost every line to the best that I can express myself.
'   But if you still have questions feel free to email me.

'Knowledge Requirements:
'   Visual Basic
'   HTML

'System Requirements:
'   Visual Basic
'   Internet Explorer 4 and above
'   SHDOCVW.dll(Microsoft Internet controls) - Components
'   Mshtml.tlb(Microsoft HTML Object library). - References
'   Msado20.tlb(Microsoft ActiveX Data Objects 2.0 Library) - References

'Author: Pinoy Ako!

Option Explicit
'Prevent you from using an undeclared varialble in the code.

Private WithEvents connConnection As ADODB.Connection
Attribute connConnection.VB_VarHelpID = -1
'connConnection is the connection with the database.

Private WithEvents rsRecordSet As ADODB.Recordset
Attribute rsRecordSet.VB_VarHelpID = -1
'rsRecordSet is the recordset that will be used with the database connection.

Private WithEvents htmlDoc As HTMLDocument
Attribute htmlDoc.VB_VarHelpID = -1
'htmlDoc is hte HTML Document that will be used with the web1(webBrowser).

Dim htmlElement As IHTMLElement
'htmlElement is the HTML Elemnt that will be used with HTMLDocument.


Private Sub htmlTempFile()
    
    Dim lCount As Integer
    'lCount integer will be used with For loop.
    
    On Error Resume Next
    'If error occur, ignore it.

    Open App.Path & "\blank.html" For Output As #1
    'Open is used to create .html file.
                  
    'The code below are mostly html code, so a little HTML knowledge is needed
    'if you want to understand them.
    'Some code are for designing purpose only, like the Javascript and CSS.
    'I'm not saying that it's a good design though. You'll be the judge for that.
                         
        Print #1, "<html>"
        'Start of html code.
        
        Print #1, "<style type=text/css>"
        Print #1, "body{scrollbar-base-color:#C6BF8E;scrollbar-arrow-color:#000000;scrollbar-track-color:#E8E0AB;}"
        Print #1, "button{width:75;color:#8A7F37;background:#FAF7E2;border:1 ouset}"
        Print #1, "a{text-decoration:none;width:130};a:hover{background:#E8E0AB};a:link{color:#8A7F37};a:visited{color:#8A7F37};a:activate{color:#8A7F37};"
        Print #1, ".tables{border:1 outset;font:12px Arial;background:#FAF7E2;width:130}"
        Print #1, "</style>"
        'CSS design of the. Includes, the Help menu,buttons and the body.
        
        Print #1, "<script language=JavaScript>"
        Print #1, "function showall() {document.all.menu2.style.visibility='Visible';}"
        Print #1, "function hidemenu() {document.all.menu2.style.visibility='hidden';}"
        Print #1, "</script>"
        'This javascript is used with the Help menu.
            
        
        Print #1, "<body style='color:firebrick;background:#FFF4F0 url(2.jpg) repeat-y;font:10pt'>"
        'The body
        
        Print #1, "<marquee  scrollamount=3 style=color#:C6BF8E;position:absolute;top:3;font:bold>Visual Basic - HTML Interface. One example is this simple ADO database.</marquee><hr>"
        Print #1, "<marquee  direction=right behavior=slide scrollamount=70 style=color:gray;position:absolute;top:25;font:italic>Author: Pinoy Ako</marquee><br><br><br>"
        'Html tag <marquee>...</mrquee>.
        'Text animation.
        
        Print #1, "<div id=menu align=center style='cursor:hand;position:absolute;top:26;left:10;width:75;height:20;"
        Print #1, "background:#FAF7E2;color:#8A7F37;border:1 outset;font:14px Arial' onMouseDown='showall()' onMouseOut='hidemenu()'>"
        Print #1, "Help</div>"
        Print #1, "<div id=menu2 align=justify style='position:absolute;top:46;left:10;visibility: Hidden ' onMouseOver='showall()' onMouseOut='hidemenu()'>"
        Print #1, "<table class=tables>"
        Print #1, "<tr><td><a id=aUsing href=#>Using the program...</a><br></td></tr>"
        Print #1, "<tr><td><a id=aAbout href=#>About...</a><br></td></tr>"
        Print #1, "</table></div>"
        'This portion is part of Help menu.
        'The html tag <a>...</a> with the ID of "aUsing" and "aAbout"
        'are used with htmlDoc_onClick function (where click events codes are found).
        
        Print #1, "<table align=right cellspacing=10 style='color:firebrick;background:url(1.jpg);font:10pt;border:C6BF8E 1 outset'> "
        Print #1, "<tr><td>First name:<form></td><td><input style=color:firebrick id=HTMLName ></td></tr>"
        Print #1, "<tr><td style=color:black>Home address:</td><td><input style=color:black id=HTMLHome ></td></tr>"
        Print #1, "<tr><td style=color:black>Email address:</td><td><input style=color:black id=HTMLEmail ></td></tr></form>"
        Print #1, "<tr><td></td><td align=center><button id=buttonAdd > Add</button> "
        Print #1, "<button id=buttonRemove > Remove</button> </td><tr>"
        'This portion is part of the  Add,Remove button and the three text box, html tags <button>...</button> and <input>.....</input>.
        'The html tags with the ID of "HTMLName","HTMLHome","HTMLEmail","buttonAdd" and "buttonRemove"
        'are used with htmlDoc_onClick function (where click events codes are found).
        
        
        Print #1, "</table><table style='font:10pt'><tr><td><ol> "
            For lCount = 0 To lstBox.ListCount - 1
                Print #1, lstBox.List(lCount) & "<br>"
                'Write the list box contents to the html file.
                'This portion is part the output of the DATABASE.
            Next lCount
        Print #1, "</td></tr></ol></table></body> </html>"
        'End of html code.
            
    Close #1
    'Close.

End Sub


Private Sub Form_Load()

    Dim strConnect As String
    'strConnect will be used as connection string.
    
    Dim strProvider As String
    'strProvider will be used as provider string.
        
    Dim strDataSource As String
    'strDataSource will be used as data source  string.
    'File location string.
              
    On Error Resume Next
    'If error occur, ignore it.
    
    strProvider = "Provider= Microsoft.Jet.OLEDB.3.51;"
    
    strDataSource = App.Path & "\data.mdb;"
    
    strDataSource = "Data Source=" & strDataSource
    
    strConnect = strProvider & strDataSource
       
    
    Set connConnection = New ADODB.Connection
    'Preparing the connection object.
    
    connConnection.CursorLocation = adUseClient
    'Client side cursor.
    'Used because this application will access the data on client machine
    'instead of server.
            
    connConnection.Open strConnect
    'Open the connection object.
      
    Set rsRecordSet = New ADODB.Recordset
    'Prepare the recordset.
    
    rsRecordSet.CursorType = adOpenStatic
    'Client side cursor type.
    
    rsRecordSet.CursorLocation = adUseClient
    'Client side cursor.
    'Used because this application will access the data on client machine
    'instead of server.
    
    rsRecordSet.LockType = adLockPessimistic
    'Insure that the record that is being edited can be saved.
    
    rsRecordSet.Source = "Select * From tblAddressBook"
    'SQL Select command.
    'Retreive the data from the table tblAddressBook.
    
    rsRecordSet.ActiveConnection = connConnection
    'It is important that the recordset know what connection to use.
    
    rsRecordSet.Open
    'Open the recordset.
    
    rsRecordSet.Sort = "fldname"
    'Sort the data in order.
         
         
     Call List
    'Load List function data.
    
    Call htmlTempFile
    'Load html function data.
     
     With Web1
        .navigate App.Path & "\blank.html"
        'Find url of the html file to be loaded in web1(webBrowser).
                
        .Move 0, 0, ScaleWidth, ScaleHeight
        'Set the web1's(webBrowser) position,height and width.
    End With
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next
    'If error occur, ignore it.
    
     With Web1
        .Move 0, 0, ScaleWidth, ScaleHeight
        'Set the web1's(webBrowser) position,height and width
        'upon resizing the application.
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Dim tempFile As String
    'tempFile will be used as pathname string
    
    tempFile = App.Path & "\blank.html"
    
    If tempFile <> "" Then Kill tempFile
    'On Unload, the html file will be deleted.
    
End Sub


Private Function htmlDoc_onclick() As Boolean
'This portion is for HTML onclick event
'and visual basic click event.

   
    Dim htmlID As String
    'htmlID will be used as the html ID string
    
    On Error Resume Next
    'If error occur, ignore it.
    
    Set htmlElement = htmlDoc.parentWindow.event.srcElement
    'Setting htmlElements equal to htmlDoc.parentWindow.event.srcElement.
    'Because, it will make the html onclick event work like visual basic
    'click event on the web1(webBrowser).
    
    htmlID = htmlElement.id
        
    Select Case htmlID
        
        Case "buttonAdd"
        '"Add" button.
        '"buttonAdd" is one of html tag <button>...</button> ID. See html code in the htmlTempFile function.
        'If you click html tag with "buttonAdd" ID, the code below will take action.
        
            rsRecordSet.Open
            'Open recordset for adding data.
            
            rsRecordSet.AddNew
            'Preparing to add new data to recordset.
                       
            rsRecordSet("fldName").Value = htmlDoc.All.Item("HTMLName").Value
            '"Name" text box.
            'The recordset fldFirstName value will be equal
            'to the value of html tag <input>..</input> with an ID of HTMLName.
            
            rsRecordSet("fldHomeAddress").Value = htmlDoc.All.Item("HTMLHome").Value
            '"Home Address" text box.
            'The recordset fldHomeAddress value will be equal
            'to the value of html tag <input>..</input> with an ID of HTMLHome.
            
            rsRecordSet("fldEmailAddress").Value = htmlDoc.All.Item("HTMLEmail").Value
            '"Email address" text box.
            'The recordset fldEmailAddress value will be equal
            'to the value of html tag <input>..</input> with an ID of HTMLEmail.
                        
            rsRecordSet.Update
            'Important to complete the adding of new data.
            'Updating the recordset.
            
            Call List
            'reload List function.
            
            Call htmlTempFile
            'reload htmlTempFile function.
            
            Web1.Refresh
            'refresh the web1(webBrowser) page.
            
        Case "buttonRemove"
        '"Remove" button.
        '"buttonRemove" is one of html tag <button>...</button> ID. See html code in the htmlTempFile function.
        'If you click html tag with "buttonRemove" ID, the code below will take action.
        
            rsRecordSet.Close
            'Close the recordset.
                       
            rsRecordSet.Open "select * from tblAddressBook where fldEmailAddress='" & htmlDoc.All.Item("HTMLEmail").Value & "'"
            'Find part of the data you wish to remove.
            
            rsRecordSet.Delete
            'Remove data.
            
            rsRecordSet.Update
            'Important to complete the deleting of data.
            'Updating the recordset.

            Call Form_Load
            'Reload Form1_Load sub to reOpen the recordset,reload List function
            'and reload htmlTempFile function.
            
            Web1.Refresh
            'refresh the web1(webBrowser) page.
            
        Case "aUsing"
        'Help menu(Using the program...)
        '"aUsing" is one of html tag <a>...</a> ID. See html code in the htmlTempFile function.
        'If you click html tag with "aUsing" ID, the code below will take action.
            
            MsgBox Space(15) & "Adding  :  Enter name, home address and email address then click Add." & Space(15) & vbNewLine & vbNewLine & _
                Space(15) & "Removing  :  Enter the email address in the ""Email address"" text box then click Remove." & Space(15) & vbNewLine & vbNewLine & _
                Space(15) & "Finding :  Press CTRL + F. " & Space(15) _
                , vbInformation, "Using the program..."
                       
        Case "aAbout"
        'Help menu(About...)
        '"aAbout" is one of html tag <a>...</a> ID. See html code in the htmlTempFile function.
        'If you click html tag with "aAbout" ID, the code below will take action.
            
            MsgBox vbNewLine & "Title : Visual Basic-HTML Interface(Ado database)." & Space(10) & vbNewLine & _
            "Author : Pinoy Ako!" & vbNewLine & "Email : marshall_hb@yahoo.com" & vbNewLine, vbInformation, "About..."
            
            
    End Select
End Function

Private Function List()
    
    Dim i As Integer
    'i integer will be used with For loop.
    
    Dim RecCnt As Integer
    'RecCnt integer will be used with recordset's Recordcount.
    
    On Error Resume Next
    'If error occur, ignore it.
    
    rsRecordSet.MoveFirst
    'move to the first recorde of database.
    
    RecCnt = rsRecordSet.RecordCount
    
    lstBox.Clear
    'Clear listbox.
    
    For i = 0 To RecCnt - 1
        
        lstBox.AddItem "<li> " & rsRecordSet("fldName") & " <br> <span style=color:black>Address: " & _
        rsRecordSet("fldHomeAddress") & "</span><br> <span style=color:black>Email:  " _
        & rsRecordSet("fldEmailAddress") & " </span><br></li>"
        'Add recordset and html string to the list box.
        'There are some html code above like the tag <li>,<br> and <span>
        'use for designing the html page.
        
                        
        rsRecordSet.MoveNext
        'Move to the next record of the database.
        
    Next i

End Function

Private Sub Web1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    
    Set htmlDoc = Web1.document
    'Setting htmlDoc equal to Web1.document is important. Because, it will make the
    'html onclick event work like visual basic click event on the web1(webBrowser) object.
   
End Sub








