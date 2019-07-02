<html>

<script language="javascript">
<!--#INCLUDE FILE="connection/connection.asp"-->

function FP_swapImg() {//v1.0
 var doc=document,args=arguments,elm,n; doc.$imgSwaps=new Array(); for(n=2; n<args.length;
 n+=2) { elm=FP_getObjectByID(args[n]); if(elm) { doc.$imgSwaps[doc.$imgSwaps.length]=elm;
 elm.$src=elm.src; elm.src=args[n+1]; } }
}

function FP_preloadImgs() {//v1.0
 var d=document,a=arguments; if(!d.FP_imgs) d.FP_imgs=new Array();
 for(var i=0; i<a.length; i++) { d.FP_imgs[i]=new Image; d.FP_imgs[i].src=a[i]; }
}

function FP_getObjectByID(id,o) {//v1.0
 var c,el,els,f,m,n; if(!o)o=document; if(o.getElementById) el=o.getElementById(id);
 else if(o.layers) c=o.layers; else if(o.all) el=o.all[id]; if(el) return el;
 if(o.id==id || o.name==id) return o; if(o.childNodes) c=o.childNodes; if(c)
 for(n=0; n<c.length; n++) { el=FP_getObjectByID(id,c[n]); if(el) return el; }
 f=o.forms; if(f) for(n=0; n<f.length; n++) { els=f[n].elements;
 for(m=0; m<els.length; m++){ el=FP_getObjectByID(id,els[n]); if(el) return el; } }
 return null;
}

function Add()
{
	document.FProgram.submit();
}

function LauchFind()
{

if (document.FSuppliers.Designation.value=='' && document.FSuppliers.IdSupplier.value=='')
	{
	alert("Renseignez au moins un champ !")
	}
	else
	{
	//window.action.('SupplierSelect_ASRaM.asp?DesignationR=' + document.FSuppliers.Designation.value + '&IdSupplierR=' + document.FSuppliers.IdSupplier.value ,'Find_Suppliers','width=650,height=300,scrollbars=yes,top=150');
	document.FSuppliers.action = "SupplierSelect_ASRaM.asp?Designation='" + document.FSuppliers.Designation.value + "'&IdSupplier='" + document.FSuppliers.IdSupplier.value + "'" ;
	document.FSuppliers.submit();
	}
}

function LauchFindADD(IdBusCase)
{//window.open("SupplierSelect_ASRaM.asp?IdSupplierDEL=' + IdBusCase");
document.FSuppliers.submit();
}


</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../css/log.css">
<title>Manage Supplier</title>
</head>

<%

IdBusCase = trim(Request.querystring("IdBusCase"))
Designation = trim(Request.querystring("Designation"))
IdSupplier = trim(Request.querystring("IdSupplier"))
IdSupplierADD = trim(Request.querystring("CaseCocher"))
delete = trim(Request.querystring("delete"))
HIdSupplier = trim(Request.querystring("HIdSupplier"))

response.write idsupplier

'affiche toute les valeurs envoyée avec la methode get 
   For Each chaine_requete In Request.QueryString
  		Response.Write "<tr>;<td>" & chaine_requete & "</td> = <td>" & Request.querystring(chaine_requete) & "</td></tr>" 
   Next

if IdSupplier <> "" or Designation <> "" then
response.write "select * from Adv_Supplier where IdSupplier like '%"&IdSupplier&"%' AND Designation like '%"& Designation & "%'"
	set GetSuppliers = CONNECTION.execute("select * from Adv_Supplier where IdSupplier like '%"&IdSupplier&"%' AND Designation like '%"& Designation & "%'")
end if

if IdBusCase <> "" then

	if IdSupplierADD <> "" then 'Sauvegarde du Frs par BC
		ps_adv_MajIndex_BC_Supplier = "ps_adv_MajIndex_BC_Supplier '" & _
		IdSupplierADD & "','" & IdBusCase & "',''"
		set MajSupplier=connection.Execute(ps_adv_MajIndex_BC_Supplier)
	end if
	if delete <> "" then 'Efface Frs
		ps_adv_MajIndex_BC_Supplier = "ps_adv_MajIndex_BC_Supplier '" & _
		HIdSupplier & "','" & IdBusCase & "','oui'"
		response.write ps_adv_MajIndex_BC_Supplier
		set MajSupplier=connection.Execute(ps_adv_MajIndex_BC_Supplier)
	end if
	
set GetExistSupplier =connection.Execute("select * from Adv_Index_BC_Supplier as BCS, Adv_Supplier as Sup" & _ 
		" where BCS.IdSupplier = Sup.IdSupplier and BCS.IdBusCase = '" & _
		IdBusCase & "' order by BCS.IdSupplier")
end if
%>
<body onunload="window.opener.location.reload()" bgcolor="#FFFFFF" text="#000000" background="../images/DefaultHomePage.gif" style="background-repeat: no-repeat;background-attachment: scroll;background-position: center center" onload="FP_preloadImgs(/*url*/'images/button1.jpg', /*url*/'images/button2.jpg')">

<form name ="FSuppliers" method="get">

<table border="0" width="100%">
	<tr>
		<td class="titrelog">Busness Case N° 

<input type="text" name="IdBusCase" value = "<%=IdBusCase%>"  readonly size="19"></td>
		<td width="58" class="titrelog"><a href="images/exit.JPG">
		<img src="images/exit_small.JPG" width="55" height="48" onclick="window.opener.location.reload(); window.close()" xthumbnail-orig-image="images/exit.JPG" border="2" ></a> </td>
	</tr>
</table>

<table border="0" width="100%">
	<tr class="titrelog">
		<td width="27%" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">Supplier 
		Id :  
		<input type="text" name="IdSupplier" onkeydown ="if (window.event.keyCode == 13) LauchFind();" size="29" maxlength="5"
		value="<%=trim(Request.querystring("IdSupplier"))%>"></td>
		<td width="34%" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6"> Supplier Name :  
		<input type="text" name="Designation" onkeydown ="if (window.event.keyCode == 13) LauchFind();" size="37" maxlength="37" 
		value="<%=trim(Request.querystring("Designation"))%>"></td>
		
		<td align="right" width="9%">
		<td align="right" width="9%">
		<img onclick= "LauchFind();" border="0" id="img10" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img10',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img10',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img10',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img10',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit">
	</tr>
</table>

&nbsp;<input type="hidden" name="ordre" value = "<%=ordre%>"  readonly size="50"><table border="0" width="100%" id="table2">
	<tr>
		<td valign="top">&nbsp;
		<b><font size="2">Select Supplier in list</font></b><table border="0" width="100%" id="table2">
		<tr> 
		<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">Supplier Code</td>
		<td width="233" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">Supplier Name </td>
		</tr>
		<%
		if IdSupplier <> "" or Designation <> "" then
			if not GetSuppliers.eof then
				do until GetSuppliers.eof
				%>
				<tr >
			
		<td width="114" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" ><%=GetSuppliers("IdSupplier")%>
		<input type="checkbox" 
		name="CaseCocher"
		value="<%=GetSuppliers("IdSupplier")%>" onclick="javascript:LauchFindADD('<%=GetSuppliers("IdSupplier")%>')">
		</td>
		<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" ><%=GetSuppliers("Designation")%>
		</td>

				</tr>
				<%
				GetSuppliers.movenext
				loop
			end if
		end if		
		%>
		
		</table>
		<p>&nbsp;</p>
		<p>&nbsp;</td>
		<td width="411" valign="top">
	
		<table border="0" width="100%" id="table3">
		<%
			if IdBusCase <> "" then
	if not GetExistSupplier.eof then
		%>
			<tr>
				<td width="97" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">Supplier Code</td>
				<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">Supplier Name</td>
				<td width="62" class="titrelog" style="padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px">&nbsp;</td>
			</tr>
	<%

		do until GetExistSupplier.eof
		%>
		<tr>
			<td width="97" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF"><%=trim(GetExistSupplier("IdSupplier"))%>&nbsp;</td>
			<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF"><%=trim(GetExistSupplier("Designation"))%>&nbsp;</td>
			<td width="60" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">

		<a href="SupplierSelect_ASRaM.asp?IdBusCase=<%=IdBusCase%>&delete='del'&Designation=<%=Designation%>&IdSupplier=<%=trim(Request.querystring("IdSupplier"))%>&HIdSupplier=<%=trim(GetExistSupplier("IdSupplier"))%>">
		<img src="images/delete.JPG" border="0" width="17" height="18">Delete</a>
		<%
		GetExistSupplier.movenext
		loop
	end if
	end if
	%>
	</table>
		</td>
	</tr>
</table>


</form>



</body>



</html>



<script language="JavaScript">
<!--#INCLUDE FILE="connection/deconnection.asp"-->
</script>