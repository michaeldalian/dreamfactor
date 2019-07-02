<html>


<script language="javascript">
<!--#INCLUDE FILE="connection/connection.asp"-->
window.opener.parent.frames("mainFrame").MainFile.submit();

function CtrlVal(Nom){// Procedure de validation
//alert(window.event.keyCode)
if (window.event.keyCode==46 || window.event.keyCode==44)
{
window.event.keyCode = 0;
} else if (window.event.keyCode<48 || window.event.keyCode>57)
{
window.event.keyCode = 0;
}}

</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="../css/log.css">
<title>Manage Current IdVariant</title>
</head>

<form  action="Variants_ASRaM.asp" name="Variants_ASRaM" method="post">

<%

IdBusCase = trim(Request("IdBusCase"))
Designation = trim(Request("Designation"))
IdVariant = trim(Request("IdVariant"))
IdVariantADD = trim(Request("CasaCocher"))
delete = trim(Request("delete"))
Status = trim(Request("Status"))
TabInd = trim(Request("TabInd"))
IdMacro = trim(Request("IdMacro"))
Var_Qty = trim(Request("Var_Qty"))
HIdVariant = trim(Request("HIdVariant"))
FIdVariant = trim(Request("FIdVariant"))

'affiche toute les valeurs envoyée avec la methode post
    for each elem in Request.Form
       response.write elem
       response.write " : "
       'response.write Request.Form(elem) 
       response.write Request(elem) 
       response.write " // "
    next


if FIdVariant <> "" or Designation <> "" or IdVersion <> "" then
	response.write "select * from Adv_Variant where IdVariant like '%"& _
	FIdVariant & "%' AND Designation like '%"& Designation & "%' AND IdVersion like '%"& IdVersion & "%'"
	 
	set RsVariants = CONNECTION.execute("select * from Adv_Variant where IdVariant like '%"& _
	FIdVariant & "%' AND Designation like '%"& Designation & "%' AND IdVersion like '%"& IdVersion & "%'")
end if

	if IdVariantADD <> "" then 'Sauvegarde des Variants
		if Var_Qty ="" then Var_Qty = 1
		ps_adv_MajIndex_BC_Variant = "ps_adv_MajIndex_BC_Variant '" & IdBusCase & "','" & IdVariantADD & "','" & _
		Status & "','" & TabInd & "','" & Var_Qty & "','" & IdMacro & "',''"
		set MajVariant=connection.Execute(ps_adv_MajIndex_BC_Variant)
		response.write ps_adv_MajIndex_BC_Variant
	elseif delete <> "" then 'efface Variant
		ps_adv_MajIndex_BC_Variant = "ps_adv_MajIndex_BC_Variant '" & IdBusCase & "','" & HIdVariant & "','" & _
		Status & "',0,0,'','" & "oui" &"'"
		set MajVariant=connection.Execute(ps_adv_MajIndex_BC_Variant)
		response.write ps_adv_MajIndex_BC_Variant
	end if
	
' Sauver vers BDD ______________________________________________________________
if Request("BSave") <> "" then
	max = trim(Request("max"))

for i=0 to max
	
	IdVariant = trim(Request("V" & i))
	Designation = trim(Request("D" & i))
	Var_Qty = trim(Request("Q" & i))
	
	if IdVariant="" and Var_Qty="" then
		else
			
		ps_adv_MajIndex_BC_Variant = "ps_adv_MajIndex_BC_Variant '" & IdBusCase & "','" & _
		IdVariant & "','" & Status & "','" & TabInd & "','" & Var_Qty & "','" & IdMacro & "',''"
		response.write "MAJ" & ps_adv_MajIndex_BC_Variant
		set MajVariant=connection.Execute(ps_adv_MajIndex_BC_Variant)
	
	end if
next
end if ' Fin Sauver______________________________________________________________

if IdBusCase <> "" then
response.write "select * from Adv_Variant, Adv_Index_BC_Variant where Adv_Variant.IdVariant=Adv_Index_BC_Variant.IdVariant AND IdBusCase like '"& IdBusCase & "%' AND Status = '"& Status & "' AND TabInd = "& TabInd & " order by Adv_Variant.IdVariant"
set GetExistVariant =connection.Execute("select * from Adv_Variant, Adv_Index_BC_Variant where Adv_Variant.IdVariant=Adv_Index_BC_Variant.IdVariant AND IdBusCase like '"& IdBusCase & "%' AND Status = '"& Status & "' AND TabInd = "& TabInd & " order by Adv_Variant.IdVariant")
end if
%>
    
    
<body bgcolor="#FFFFFF" text="#000000" background="../images/DefaultHomePage.gif" style="background-repeat: no-repeat;background-attachment: scroll;background-position: center center">

<table border="0" width="100%">
	<tr>
		<td class="titrelog"><table border="1" width="83%" id="table2">
	<tr>
		<td width="251">Busness Case N° 
		<input type="text" name="IdBusCase" value = "<%=IdBusCase%>"  readonly size="19"></td>
		<td>Status
		<input type="text" name="Status" value = "<%=Status%>"  readonly size="19"></td>
		<td>Tab Index <input type="text" name="TabInd" value = "<%=TabInd%>"  readonly size="19"> </td>
	</tr>
</table> 
		
		</td>
		<td width="58" class="titrelog"><a href="images/exit.JPG">
		<img src="images/exit_small.JPG" width="55" height="48" onclick="window.close()" xthumbnail-orig-image="images/exit.JPG" border="2" style="border: 1px solid #FF0000" ></a> 
		</td>
	</tr>
</table>

<table border="0" width="100%">
<tr class="rublog" >
		<td width="309" >
		
		<table border="0" width="100%">
<tr class="rublog" >
		<td width="19%" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		Variant <input type="text" name="FIdVariant" size="14" maxlength="14" value="<%=Request("FIdVariant")%>"></td>
		<td width="60%" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6"> 
		Variant Description :  <input type="text" name="designation" size="27" maxlength="27" value="<%=Request("designation")%>"></td>
		<td width="34%" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6"> 
		Version : <input type="text" name="IdVersion" size="37" maxlength="37" value="<%=trim(Request("IdVersion"))%>"></td>
<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" >
<input type="submit" value=" Find " name="BFind"></td>
</tr> 
</table> 
<%
if Request("designation") <> "" or Request("FIdVariant") <> "" then
if not RsVariants.eof then
%>
<table border="1" width="100%" id="table1" cellspacing="0" cellpadding="0" bordercolordark="#000000">
	<tr class=titrelog  align="center" bgcolor="#DDDDFF">
		<td bgcolor="#DCE9F6">Variant</td>
		<td width="155" bgcolor="#DCE9F6">IdVariant Description</td>
	</tr>
		<%	

	do until RsVariants.eof
	flush = flush + 1
		
	response.write "<tr class=rublog >" & _
		"<td>"&RsVariants("IdVariant")&"<input type=""checkbox""name=""casacocher""	value="&"'"&RsVariants("IdVariant")&"'"&" onclick=""submit()""></td>" & _
		"<td>&nbsp;<A class=""valider"" href=""#"" onclick=""Submit()"">"&RsVariants("designation")&"</A></td>" & _
	"</tr>" 
			
			if flush = 800 then
			response.flush
			flush = 0
			end if
	RsVariants.movenext
	loop

	%>
	
</table>
<% end if 
end if
%>
		
		
	</td>
	<td valign="top">
		

		<table border="0" width="100%">
			<tr>
				
				<td bgcolor="#DCE9F6" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog">
				Current Variant</td>
				<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#DCE9F6">Designation</td>
				<td width="43" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#DCE9F6">
				Qty</td>
				<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#DCE9F6">
				<input type="submit" value="Save" name="BSave" style="color:#0000FF"></td>
			</tr>
					<%
		if not GetExistVariant.eof then
			i=1
			Response.write "IdMacro <input type=""text"" name=""IdMacro"" size=""14"" value="&GetExistVariant("IdMacro")&" >"
			do until GetExistVariant.eof
			%>
			<tr>
	
				<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
				<input type="text" name="V<%=i%>" value="<%=GetExistVariant("IdVariant")%>" readonly ></td>
				<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
				<input type="text" readonly name="D<%=i%>" value="<%=trim(GetExistVariant("Designation"))%>"></td>
				<td width="43" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
				<input type="text" name="Q<%=i%>" value="<%=trim(GetExistVariant("Var_Qty"))%>" onkeypress ="javascript:CtrlVal(this)" size="3"></td>
				<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
				<a href="Variants_ASRaM.asp?IdBusCase=<%=IdBusCase%>&TabInd=<%=TabInd%>&Status=<%=Status%>&FIdVariant=<%=FIdVariant%>&Delete=<%=i%>&HIdVariant=<%=trim(GetExistVariant("IdVariant"))%>" >Delete</a>
				</td>
				
			</tr>
			<%
			cpt = cint(i)
			i=i+1
			GetExistVariant.movenext
		loop
		
		 end if
		if casacocher <> "" then
		cpt = cpt + 1
		%>
		<input type="hidden" name="N<%=cpt%>" value="<%=cpt%>" size="2" readonly >
		<%
		end if
 %>		
		<input type="hidden" name="max" value="<%=cpt+1%>" size="20">
	</table>

		
		</td>

</tr> 
</table>


</body>
</form>

<script language="JavaScript">
<!--#INCLUDE FILE="connection/deconnection.asp"-->
</script>

</html>