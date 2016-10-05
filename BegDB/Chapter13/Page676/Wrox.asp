<HTML>
<TITLE>Database Programming with Visual Basic 6.0</TITLE></HEAD>

<CENTER>
<H1><font size=4>Using ADO in an Active Server Page</H1></font>
<H2>Database Programming with Visual Basic 6.0</H2><br>

<%
dim myConnection
dim connectString

connectString = "Provider=Microsoft.Jet.OLEDB.3.51;Data Source=C:\begdb\biblio.mdb"

Set myConnection = Server.CreateObject("ADODB.Connection")
Set RSTitleList = Server.CreateObject("ADODB.Recordset")

myConnection.Open connectString

Set RSTitleList =  myConnection.Execute( "Select * From titles WHERE PubID = 42") %>

<TABLE align=center COLSPAN=8 CELLPADDING=5 BORDER=0 WIDTH=200>

<!-- Begin our column header row -->  
<TR>
   <TD  VALIGN=TOP BGCOLOR="#800000">
     <FONT STYLE="ARIAL NARROW" COLOR="#ffffff" SIZE=2>      Publisher 
      ID</FONT> 
   </TD>
   <TD ALIGN=CENTER BGCOLOR="#800000">
     <FONT STYLE="ARIAL NARROW" COLOR="#ffffff" SIZE=2>      Title
     </FONT>
   </TD>
</TR>

<!-- Ok, let's get our data now -->
<% do while not RStitleList.EOF %>   <TR>
   <TD BGcolor ="f7efde" align=center><font style ="arial narrow" size=2>
           <%=RStitleList("PubID")%></font>   </TD>

   <TD BGcolor ="f7efde" align=center><font style ="arial narrow" size=2>
          <%=RSTitleList("Title") %>   </font>   </TD>   </TR>

   <% RSTitleList.MoveNext%>
<%loop %>

</TABLE>
</CENTER>
</BODY>
</HTML>
