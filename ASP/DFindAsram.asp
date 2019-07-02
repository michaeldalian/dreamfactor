<html>

<script language="javascript">
<!--#INCLUDE FILE="connection/connection.asp"-->


function MainASRaM(IdBusCase)
	{
	window.opener.parent.frames("mainFrame").location.href="BusEdit_ASRaM.asp?IdBusCase=" + IdBusCase + "&Find=Find" ;
	self.close();
	}


</script>
<script language="javascript" src="../images/function.js">
</script>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Result of your Search</title>
<link rel="stylesheet" type="text/css" href="../css/log.css">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../images/DefaultHomePage.gif" style="background-repeat: no-repeat;background-attachment: scroll;background-position: center center">

<%

'affiche toute les valeurs envoyée avec la methode get 
   For Each chaine_requete In Request.QueryString
  		Response.Write "<tr>;<td>" & chaine_requete & "</td> = <td>" & Request.QueryString(chaine_requete) & "</td></tr>" 
   Next

' Controle si le user existe


%>
<a  class=valider href="../excel/excel.asp?requete=<%=r_xls%>">
	<img class=valider border=0 src="../images/iconeexcel.jpg"> Export to Excel 
</a>
<%

IdBusCase = request.querystring("IdBusCase")
IdBuyer = request.querystring("IdBuyer")
IdSupplier = request.querystring("IdSupplier")
Title = request.querystring("Title")
IdALStation = request.querystring("IdALStation")

if IdSupplier <> "" then
SQLFind= "select * from Adv_BusCase as BC, Adv_Index_BC_Supplier as Sup where" & _
" BC.IdBusCase = Sup.IdBusCase and IdSupplier like '%" & IdSupplier & "%' and"
else
SQLFind= "select * from Adv_BusCase as BC where"
end if

SQLFind= SQLFind & _
" BC.IdBusCase like '%" & IdBusCase & "%'" &_
" and IdBuyer like '%" & IdBuyer & "%'" & _
" and Title like '%" & Title & "%'" & _
" and IdALStation like '%" & IdALStation & "%'" & _
" order by BC.IdBusCase"




response.write SQLFind
set GetASRaM = CONNECTION.execute(SQLFind)

'
	dim nb
	if not GetASRaM.eof then
		do until GetASRaM.eof
			nb = nb + 1
			GetASRaM.movenext
		loop
		GetASRaM.movefirst
    end if
    
    response.write nb
    if nb = 1 then 
response.write nb
    %>
    
      <script language="javascript">
		window.opener.parent.frames("mainFrame").location.href="BusEdit_ASRaM.asp?IdBusCase=" +'<%=GetASRaM("IdBusCase")%>' + "&Find=Find";
		setTimeout("self.close();",1);
		</script>

	<%
	end if
%>

<form name="find" method =get>
<p align="center" class=titrelog><i>Result of your Search </i> </p>
<table border="1" width="100%" id="table1" cellspacing="0" cellpadding="0" bordercolordark="#000000">
	<tr class=titrelog>
		<td align="center" width="110" bgcolor="#DCE9F6" >Business Case N°</td>
		<td align="center" bgcolor="#DCE9F6" width="219" >Title</td>
		<td align="center" bgcolor="#DCE9F6" >Creator Name</td>
		<td align="center" bgcolor="#DCE9F6" >Buyer Name</td>
		<!--<td align="center" bgcolor="#DCE9F6" >Team</td>-->
		<td align="center" bgcolor="#DCE9F6" >Supplier Name</td>
		<td align="center" bgcolor="#DCE9F6" width="107" >Assembly Line</td>
	</tr >
	<%	
	if not GetASRaM.eof then
	
	do until GetASRaM.eof
		%>
		<tr class=rublog onmouseover="mOvr(this,'#DCE9F6');" onmouseout="mOut(this,'');">
		<td width="110" ><img src="../images/vide.gif" width="15" height="1"><A class="valider" href="#" onclick="javascript:MainASRaM('<%=GetASRaM("IdBusCase")%>')"><%=GetASRaM("IdBusCase")%></A></td>
		<td width="219" ><img src="../images/vide.gif" width="15" height="1"><A class="valider" href="#" onclick="javascript:MainASRaM('<%=GetASRaM("IdBusCase")%>')"><%=GetASRaM("Title")%></A></td>
		<td ><img src="../images/vide.gif" width="15" height="1"><A class="valider" href="#" onclick="javascript:MainASRaM('<%=GetASRaM("IdBusCase")%>')"><%=GetASRaM("IdCreator")%></A></td>
		<td ><img src="../images/vide.gif" width="15" height="1"><A class="valider" href="#" onclick="javascript:MainASRaM('<%=GetASRaM("IdBusCase")%>')"><%=GetASRaM("IdBuyer")%></A></td>
		<td ><img src="../images/vide.gif" width="15" height="1"><A class="valider" href="#" onclick="javascript:MainASRaM('<%=GetASRaM("IdBusCase")%>')">
		<%if IdSupplier <> "" then 
			response.write "<"&char(37)&"=GetASRaM(""IdSupplier"")"&char(37)&"></A></td>"
		end if%></A></td>
		<td ><img src="../images/vide.gif" width="15" height="1"><A class="valider" href="#" onclick="javascript:MainASRaM('<%=GetASRaM("IdBusCase")%>')"><%=GetASRaM("IdALStation")%></A></td>
<%	response.write "</tr>"
	GetASRaM.movenext
	loop
   end if
	%>
	
</table>


&nbsp;<p><A class="valider" href="#" onClick = "javascript:window.close();">Close</A>  <A onClick = "javascript:window.close();" href="#">
<IMG  border=0 height=16 hspace=5 src="images/Create.JPG" width=16 ></a></td>
</form>
</body>
<script language="JavaScript">
<!--#INCLUDE FILE="connection/deconnection.asp"-->
</script>
</p>

</html>