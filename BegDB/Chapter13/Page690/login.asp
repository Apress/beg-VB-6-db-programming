<!-- Notice that the cookie code is before the HTML tag -->
<%
dim cookieValue
cookieValue = Request.Cookies("visitorNumber") 
%>

<!-- Now we start the actual page -->
<HTML>
<HEAD>
<TITLE>Programming databases with VB6 Cookie Example</TITLE>
</HEAD>
<BODY>
<CENTER>
<H1>Welcome to the ADO Cookie Web Site</H1>
<HR>


<FORM NAME="login" Action="visitor.asp" method="POST">
   <INPUT TYPE="hidden" NAME="cookieValue" VALUE="<%=cookieValue%>"><P>
   <P>Press Enter to log in to my web site</P>
   <P><INPUT TYPE="submit" value="Log In"> 
</FORM>
</CENTER>
<HR>
<H5>Copyright: Programming Databases with Visual Basic 6.0.<BR>
Last revised: July 19, 1998</h5>
</BODY>
</HTML>
