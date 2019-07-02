<html>


<script language="javascript">
<!--#INCLUDE FILE="connection/connection.asp"-->

function OuvrirFenetre(URL,Nom,Param)
{
if (URL==''||URL==undefined)
	{
	alert("Renseignez au moins un champ !")
}else{

	if (URL.substr(URL.length-4, 3)=='asp')
		{
		URL= URL + "IdBusCase=" + document.getElementById('IdBusCase').value
	}else{
		URL= URL + "&IdBusCase=" + document.getElementById('IdBusCase').value
	}
	if (Nom!=undefined&&Nom!='')
		{Nom=Nom.replace(' ','_')
		Nom=Nom.replace(' ','_')
		Nom=Nom.replace(' ','_')
		Nom=Nom.replace(' ','_')
	}

	if (Param!=undefined&&Param!='')
		{Param=Param + ",scrollbars,width=" +(screen.width-200)+ ",height=" +(screen.height-420)+ ",top=50,left=50,directories,location,resizable"
	}else{
		Param="scrollbars,width=" +(screen.width-200)+ ",height=" +(screen.height-420)+ ",top=50,left=50,directories,location,resizable"
	}
	window.open(URL,Nom,Param) ;
}}


</script>

<%
set RsAssemblyLine =connection.execute("select * from Adv_AssemblyLine order by IdALStation")
set RsBuyer =connection.execute("select * from Adv_Buyer order by IdBuyer")
set RsBusCase =connection.execute("select * from Adv_BusCase where IdBusCase = '" & IdBusCase & "' order by IdBusCase")
%>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../css/log.css">

</head>

<body bgcolor="#FFFFFF" text="#000000" background="../images/DefaultHomePage.gif" style="background-repeat: no-repeat;background-attachment: scroll;background-position: center center">

<form name="FindAsram" onkeydown ="if (window.event.keyCode == 13) OuvrirFenetre('DFindAsram.asp?','Result','target=\'mainFrame\'');">
<table border="0" width="100%" id="table1" cellpadding="6" cellspacing="1">
	<tr>
		<td width="178" class="titrelog">&nbsp;</td>
		<td width="412">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td width="178" class="titrelog">&nbsp;</td>
		<td class="titrelog" width="412"><font size="5">Find Existing Business Case</font></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td width="178" class="titrelog">&nbsp;</td>
		<td width="412">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td width="178" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		Business Case N° &nbsp;</td>
		<td width="412" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF"> 
		<!-- Cree un textBox avec contrôle de valeur -->
		<input type="text" name="IdBusCase" value="<%=Request.QueryString("IdBusCase")%>" size="20"></p>
</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td width="178" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		Business Case Title</td>
		<td width="412" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF"> 
		<input type="text" name="Title" size="20" value=<%=Request.QueryString("Title")%>></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td width="178" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		Creator Name</td>
		<td width="412" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">
		

		
		<input type="text" name="IdCreator" size="20" value=<%=Request.QueryString("IdCreator")%>></td>
		<td>&nbsp;</td>
		<td>
		&nbsp;</td>
	</tr>

	<tr>
		<td width="178" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">Buyer Name</td>
		<td width="412" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">

	<select size="1" name="IdBuyer" >
	<OPTION></OPTION>
	<%RsBuyer.movefirst
	do until RsBuyer.eof
		if IdBuyer = trim(RsBuyer("IdBuyer")) then %>
			<OPTION Value ="<%=trim(RsBuyer("IdBuyer"))%>" selected >
			<%=RsBuyer("Designation")%></OPTION>
		<%else%>
			<OPTION Value ="<%=trim(RsBuyer("IdBuyer"))%>">
			<%=RsBuyer("Designation")%></OPTION>
		<%end if
		RsBuyer.movenext
	loop%>
	</SELECT></td>
		<td align="right" width="52">&nbsp;</td>		
		<td><A onClick = "OuvrirFenetre('DFindAsram.asp?','Result','target=\'mainFrame\'');" href="#">
		<IMG  border=0 height=17 hspace=5 src="images/Find1.jpg" width=85 ></a></td>
	</tr>
	<tr>
		<td width="178" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		Supplier Code</td>
		<td width="412" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">

				
		<input type="text" name="IdSupplier" size="20" value=<%=Request.QueryString("IdSupplier")%>></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td width="178" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		Assembly Line</td>
		<td width="412" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">

		
		
	<select size="1" name="IdALStation" >
	<OPTION></OPTION>
	<%RsAssemblyLine.movefirst
	do until RsAssemblyLine.eof
		if IdALStation = trim(RsAssemblyLine("IdALStation")) then %>
			<OPTION Value ="<%=trim(RsAssemblyLine("IdALStation"))%>" selected >
			<%=RsAssemblyLine("Designation")%></OPTION>
		<%else%>
			<OPTION Value ="<%=trim(RsAssemblyLine("IdALStation"))%>">
			<%=RsAssemblyLine("Designation")%></OPTION>
		<%end if
		RsAssemblyLine.movenext
	loop%>
	</SELECT></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	<tr>
		<td width="178" class="titrelog" height="33">&nbsp;</td>
		<td width="412" height="33">&nbsp;</td>
		<td height="33">&nbsp;</td>
		<td height="33">&nbsp;</td>
	</tr>
	
	<tr>
		<td > </td>
		<td > <input type="hidden" name="Team_member" size="40"></td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
	
	<tr>
		<td width="178" class="titrelog">&nbsp;</td>
		<td width="412">&nbsp;</td>
		<td>&nbsp;</td>
		<td>&nbsp;</td>
	</tr>
</table>
</form>

</body>


<script language="JavaScript">
<!--#INCLUDE FILE="connection/deconnection.asp"-->
</script>

</html>