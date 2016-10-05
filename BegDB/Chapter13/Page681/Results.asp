<HTML>
<HEAD>
<TITLE>Database Programming with Visual Basic 6.0</TITLE>
</HEAD>

<CENTER>
<H1><FONT size=4>Requesting Publisher's Titles</H1></FONT>
<H2>Database Programming with Visual Basic 6.0</H2><BR>


<%

dim myConnection   
dim rsTitleList
dim connectString
dim sqlString
dim requestPubID

connectString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=C:\begdb\biblio.mdb"

Set myConnection = Server.CreateObject("ADODB.Connection")
Set rsTitleList = Server.CreateObject("ADODB.Recordset")

myConnection.Open connectString

requestPubID = Request.Form("PubID")

sqlString = "Select * From titles WHERE PubID = " & requestPubID

Set RSTitleList =  myConnection.Execute(sqlString) 

If (RSTitleList.BOF) AND (RSTitleList.EOF) then
  Response.Write("Sorry, but Publisher Number " & requestPubID & " was not found.")
ELSE
%>

<TABLE align=center COLSPAN=8 CELLPADDING=5 BORDER=0 WIDTH=200>
<!-- BEGIN column header row -->  
<TR>
   <TD  VALIGN=TOP BGCOLOR="#800000">
     <FONT STYLE="ARIAL NARROW" COLOR="#ffffff" SIZE=2>
        Publisher ID
     </FONT>
   </TD>
   <TD ALIGN=CENTER BGCOLOR="#800000">
     <FONT STYLE="ARIAL NARROW" COLOR="#ffffff" SIZE=2>
       Title
     </FONT>
   </TD>
</TR>
<!-- Get Data -->
<% do while not RStitleList.EOF %>
<TR>
   <TD BGcolor ="f7efde" align=center>
     <font style ="arial narrow" size=2>
        <%=RStitleList("PubID")%>
     </font>
   </TD>
   <TD BGcolor ="f7efde" align=center>
     <font style ="arial narrow" size=2>
       <%=RSTitleList("Title") %>
     </font>
   </TD>
</TR>
<% RSTitleList.MoveNext%>
<%loop %><!-- Next Row -->
</TABLE>
</center>
</BODY>
<% End if %>
</HTML>
<HTML>
