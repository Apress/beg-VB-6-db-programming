<%
visitorCookie = Request.Form("cookieValue")
if visitorCookie = "" then
  response.Redirect "newUser.asp"
End If
%>

<HTML>
<HEAD>
<TITLE>Database Programming with Visual Basic 6.0</TITLE>
</HEAD>

<BODY>
<CENTER>
<H1><font size=4>Retrieving Visitor Information</font></H1>
<H2>Database Programming with Visual Basic 6.0</H2>
</CENTER>
<BR>
<B>

<%
Dim myDll
Dim myArray 
dim firstName
dim lastName
dim previousVisit
dim totalVisits
dim secondsAgo

Set myDll = Server.CreateObject("trackVisitors.visitors")

myArray = myDll.getvisitor(visitorCookie)

firstName = myArray(0)
lastName = myArray(1)
previousVisit = myArray(2)
totalVisits = myArray(3)

Response.Write("Welcome back ")
Response.Write(firstName)
Response.Write(" ")
Response.Write(lastName)
%><P>
<%
Response.Write("You have visited my web site: ")
Response.Write(totalVisits)
Response.Write(" times.") 
%><P>
<%
Response.Write("The last time you were here was ")
Response.Write(previousVisit)
%>
<P>

</B>
<HR>
<h5>Copyright: Programming Databases with Visual Basic 6.0.<br>
Last revised: May 24, 1998</h5>
</BODY>
</HTML>
