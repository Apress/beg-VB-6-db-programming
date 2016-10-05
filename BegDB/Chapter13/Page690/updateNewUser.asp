<%
Dim myDll
Dim myArray 
Dim cookieID
dim firstName
dim lastName

firstName = Request.Form("firstName")
lastName = Request.Form("lastName")

Set myDll = Server.CreateObject("trackVisitors.visitors")

cookieID = myDll.setvisitor(firstName, lastName)

Response.Cookies("visitorNumber") = cookieID
Response.Cookies("visitorNumber").Expires = "December 30, 2000"

%>

<HTML>
<TITLE>Database Programming with Visual Basic 6.0</TITLE></HEAD>

<CENTER>
<H1><font size=4>Updating New User</font></H1>
<H2>Database Programming with Visual Basic 6.0</H2><BR>
</CENTER>
<B>
Welcome to my site
  <% Response.write(firstName)
     Response.Write(" ")
     Response.Write(lastName) %> 
     <P>
     <%
     Response.Write(Request.Cookies("visitorNumber"))
     Response.Write (" is the cookie just written to your system.")
  %>
<P>
This is your first visit on <%=now %>
</B>
<HR>
<h5>Copyright: Programming Databases with Visual Basic 6.0.<BR>
Last revised: May 24, 1998</H5>
</HTML >
