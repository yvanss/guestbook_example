<%
'****************************************************************************************
'**  Copyright Notice    
'**
'**  WWeb Wiz Guide Database Tutorial
'**                                                              
'**  Copyright 2001-2002 Bruce Corkhill All Rights Reserved.                                
'**
'**  This program is free software; you can modify (at your own risk) any part of it 
'**  under the terms of the License that accompanies this software and use it both 
'**  privately and commercially.
'**
'**  All copyright notices must remain in tacked in the scripts and the 
'**  outputted HTML.
'**
'**  You may use parts of this program in your own private work, but you may NOT
'**  redistribute, repackage, or sell the whole or any part of this program even 
'**  if it is modified or reverse engineered in whole or in part without express 
'**  permission from the author.
'**
'**  You may not pass the whole or any part of this application off as your own work.
'**   
'**  All links to Web Wiz Guide and powered by logo's must remain unchanged and in place
'**  and must remain visible when the pages are viewed unless permission is first granted
'**  by the copyright holder.
'**
'**  This program is distributed in the hope that it will be useful,
'**  but WITHOUT ANY WARRANTY; without even the implied warranty of
'**  MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE OR ANY OTHER 
'**  WARRANTIES WHETHER EXPRESSED OR IMPLIED.
'**
'**  You should have received a copy of the License along with this program; 
'**  if not, write to:- Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom.
'**    
'**
'**  No official support is available for this program but you may post support questions at: -
'**  http://www.webwizguide.info/forum
'**
'**  Support questions are NOT answered by e-mail ever!
'**
'**  For correspondence or non support questions contact: -
'**  info@webwizguide.com
'**
'**  or at: -
'**
'**  Web Wiz Guide, PO Box 4982, Bournemouth, BH8 8XP, United Kingdom
'**
'****************************************************************************************


'Dimension variables
Dim adoCon 			'Holds the Database Connection Object
Dim rsAddComments		'Holds the recordset for the new record to be added to the database
Dim strSQL			'Holds the SQL query for the database

'Create an ADO connection odject
Set adoCon = Server.CreateObject("ADODB.Connection")

'Set an active connection to the Connection object using a DSN-less connection
adoCon.Open "DRIVER={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("guestbook.mdb")

'Set an active connection to the Connection object using DSN connection
'adoCon.Open "DSN=guestbook"

'Create an ADO recordset object
Set rsAddComments = Server.CreateObject("ADODB.Recordset")

'Initialise the strSQL variable with an SQL statement to query the database
strSQL = "SELECT tblComments.Name, tblComments.Comments FROM tblComments;"

'Set the cursor type we are using so we can navigate through the recordset
rsAddComments.CursorType = 2

'Set the lock type so that the record is locked by ADO when it is updated
rsAddComments.LockType = 3

'Open the tblComments table using the SQL query held in the strSQL varaiable
rsAddComments.Open strSQL, adoCon

'Tell the recordset we are adding a new record to it
rsAddComments.AddNew

'Add a new record to the recordset
rsAddComments.Fields("Name") = Request.Form("name")
rsAddComments.Fields("Comments") = Request.Form("comments")

'Write the updated recordset to the database
rsAddComments.Update

'Reset server objects
rsAddComments.Close
Set rsAddComments = Nothing
Set adoCon = Nothing

'Redirect to the guestbook.asp page
Response.Redirect "guestbook.asp"
%>