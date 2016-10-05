<HTML>
<HEAD><TITLE>Our First ASP Script</TITLE></HEAD>
<H1><B>Our First ASP Script</B></H1>
<BODY>

Let's count up to 5.

<BR>
<HR>
<% For iCounter = 1 to 5  
  Response.Write(iCounter) %>
<BR>
<% Next %>
<HR>
</BODY>
</HTML>
