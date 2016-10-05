<HTML>
<HEAD>
<TITLE>Database Programming with Visual Basic 6.0</TITLE>
</HEAD>
<BODY>
<CENTER>
<H1><FONT size=4>
Using ADO in a Visual Basic Script Web Page
</FONT></H1>
<H2>Database Programming with Visual Basic 6.0</H2>
<HR>

<! Begin server side script here>

<%

dim myconnection
dim rsTitleList

set myConnection = Server.CreateObject("ADODB.Connection")

myconnection.open "Provider=Microsoft.Jet.OLEDB.3.51;" _
                 & "Data Source=C:\begdb\biblio.mdb"
  
SQLQuery = "SELECT title FROM titles"

set rsTitleList =  myConnection.Execute(SQLQuery)

do until rsTitleList.eof
  Response.Write rsTitleList("Title")  %>
  <BR>
  <%
  rsTitleList.movenext
loop

rsTitleList.close
set rsTitleList = nothing
%>
<! end server side script>
<HR>

</CENTER>
</BODY>
</HTML>
