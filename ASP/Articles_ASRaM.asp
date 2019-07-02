<html>


<script language="javascript">
<!--#INCLUDE FILE="connection/connection.asp"-->
window.opener.parent.frames("mainFrame").MainFile.submit();
</script>
<script language="javascript">
function OldMat(IdArticle)
	{
	document.Articles_ASRaM.submit();

	}
function CtrlVal(Nom){// Procedure de validation
//alert(window.event.keyCode)
if (window.event.keyCode==46 || window.event.keyCode==44)
{
window.event.keyCode = 44;
} else if (window.event.keyCode<48 || window.event.keyCode>57)
{
window.event.keyCode = 0;
}}

</script>

<%

IdBusCase = trim(Request("IdBusCase"))
Designation = trim(Request("Designation"))
IdArticle = trim(Request("IdArticle"))
IdArticleADD = trim(Request("CasaCocher"))
delete = trim(Request("delete"))
Status = trim(Request("Status"))
TabInd = trim(Request("TabInd"))
HIdArticle = trim(Request("HIdArticle"))
FIdArticle = trim(Request("FIdArticle"))
FDesignation = trim(Request("FDesignation"))
IdLine = trim(Request("IdLine"))
if Request("HidLine")<>"" then HidLine = Clng(Request("HidLine"))
ValidTo = trim(Request("ValidTo"))
IdCurrency = trim(Request("IdCurrency"))

'affiche toute les valeurs envoyée avec la methode post
    for each elem in Request.Form
       response.write elem
       response.write " : "
       'response.write Request.Form(elem) 
       response.write Request(elem) 
       response.write " // "
    next

	if IdArticleADD <> "" then 'Ajout d'un Article
		HidLine=HidLine+1
		set RsTempArt=connection.Execute("SELECT DISTINCT IdArticle, Art_Price AS Art_Price, Art_Qty, IdUnit FROM dbo.Adv_Index_BC_Article ABC WHERE IdBusCase = '"& IdBusCase & "' AND Status = '"& Status & "' AND ValidTo = 9999 AND IdArticle = '"&IdArticleADD&"' AND Art_Price <> 0")
		if not(RsTempArt.eof) then
			NewArt_Price=RsTempArt("Art_Price")
			NewArt_Qty=RsTempArt("Art_Qty")
			NewArt_IdUnit=RsTempArt("IdUnit")
		else
			if NewArt_Price="" then NewArt_Price=0
			if NewArt_Qty="" then NewArt_Qty=0
			if NewArt_IdUnit="" then NewArt_IdUnit="PCE"
		end if

		ps_adv_MajIndex_BC_Article = "ps_adv_MajIndex_BC_Article '"&IdBusCase&"'," & _
		HidLine&",'"&IdArticleADD&"',"&TabInd&",'"&Status&"',"&9999&",'"&NewArt_Qty&"','"&NewArt_Price&"','"&NewArt_IdUnit&"','"& _
		IdCurrency&"',''"
		response.write ps_adv_MajIndex_BC_Article
		set MajArticle=connection.Execute(ps_adv_MajIndex_BC_Article)
	elseif delete <> "" and Request("BFind")<>" Find " then 'efface IdArticle
		ps_adv_MajIndex_BC_Article = "ps_adv_MajIndex_BC_Article '"&IdBusCase&"'," & _
		IdLine&",'',"&TabInd&",'"&Status&"',"&ValidTo&",'','','','','oui'"
		response.write ps_adv_MajIndex_BC_Article
		set MajArticle=connection.Execute(ps_adv_MajIndex_BC_Article)
	end if
	
' Sauver vers BDD ______________________________________________________________
if Request("BSave") <> "" then
	RsArtCount = trim(Request("RsArtCount"))
	if RsArtCount = "" then%>
		<script language="javascript">
			this.window.close()
		</script>
	<%end if
	
	for i=0 to RsArtCount
		
		IdArticle = trim(Request("A" & i))
		Designation = trim(Request("D" & i))
		Art_Qty = replace (trim(Request("Q" & i)),",",".")
		Art_Price = replace (trim(Request("P" & i)),",",".")
		IdUnit = trim(Request("U" & i))
		ValidTo = trim(Request("V" & i))
		IdLine=trim(Request("L" & i))
	
		if not IdArticle="" then
			ps_adv_MajIndex_BC_Article = "ps_adv_MajIndex_BC_Article '"&IdBusCase&"'," & _
			IdLine&",'"&IdArticle&"',"&TabInd&",'"&Status&"',"&ValidTo&","& _
			Art_Qty&","&Art_Price&",'"&IdUnit&"','"&IdCurrency&"',''"
			response.write "MAJ" & ps_adv_MajIndex_BC_Article
			set MajArticle=connection.Execute(ps_adv_MajIndex_BC_Article)
		end if
	next
end if ' Fin Sauver______________________________________________________________

if FIdArticle <> "" or FDesignation <> "" then
	set GetArticles = CONNECTION.execute("select * from Adv_Article where IdArticle like '%"& _
	FIdArticle & "%' AND Designation like '%"&FDesignation&"%'" & _
	" AND IdArticle not in ( select IdArticle from Adv_Index_BC_Article where IdBusCase = '"& _
	IdBusCase&"' AND TabInd = "&TabInd&" AND Status = '"&Status &"')")

response.write "FIND select * from Adv_Article where IdArticle like '%"& _
	FIdArticle & "%' AND Designation like '%"&FDesignation&"%'" & _
	" AND IdArticle not in ( select IdArticle from Adv_Index_BC_Article where IdBusCase = '"& _
	IdBusCase&"' AND TabInd = "&TabInd&" AND Status = '"&Status &"')"

end if

set GetExistArticle =connection.Execute("select * from Adv_Article as A, Adv_Index_BC_Article as ABC where A.IdArticle=ABC.IdArticle AND IdBusCase = '"& IdBusCase & "' AND TabInd = "& TabInd & " AND ValidTo=9999 AND Status = '"& Status & "' order by ABC.IdLine, A.IdArticle")
response.write "SELL select * from Adv_Article as A, Adv_Index_BC_Article as ABC where A.IdArticle=ABC.IdArticle AND IdBusCase = '"& IdBusCase & "' AND TabInd = "& TabInd & " AND ValidTo=9999 AND Status = '"& Status & "' order by ABC.IdLine, A.IdArticle"
set RsCurrency =connection.execute("select * from Adv_Currency")
set RsUnit =connection.execute("select * from Adv_Unit")

%>
    
    
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<link rel="stylesheet" type="text/css" href="http://intra-dev/intrapur/procurement_mrs/css/log.css">
<title>Manage Current IdArticle</title>
</head>

<form action="Articles_ASRaM.asp" name="Articles_ASRaM" method="post" onkeydown ="if (window.event.keyCode == 13) {document.getElementById('BSave').click();}">
<body bgcolor="#FFFFFF" text="#000000" background="http://intra-dev/intrapur/procurement_mrs/images/DefaultHomePage.gif" style="background-repeat: no-repeat;background-attachment: scroll;background-position: center center">

<table border="0" width="100%">
	<tr>
		<td class="titrelog"><table border="1" width="100%" id="table2">
	<tr>
		<td style="font-size: 10pt" width="245">Busness Case N° 
		<input type="text" name="IdBusCase" value = "<%=IdBusCase%>" readonly size="19"></td>
		<td style="font-size: 10pt">Status
		<input type="text" name="Status" value = "<%=Status%>"  readonly size="19"></td>
		<td style="font-size: 10pt">Tab N° <input type="text" name="TabInd" value = "<%=TabInd%>"  readonly size="19"> </td>
		<td style="font-size: 10pt">Currency: 
			<select size="1" name="IdCurrency" >
				<OPTION></OPTION>
				<%RsCurrency.movefirst
				do until RsCurrency.eof
					if IdCurrency = trim(RsCurrency("IdCurrency")) then %>
						<OPTION Value ="<%=trim(RsCurrency("IdCurrency"))%>" selected >
						<%=RsCurrency("Designation")%></OPTION>
					<%else%>
						<OPTION Value ="<%=trim(RsCurrency("IdCurrency"))%>">
						<%=RsCurrency("Designation")%></OPTION>
					<%end if
					RsCurrency.movenext
				loop%>
			</SELECT>
		</td>
	</tr>
</table> 
		
		</td>
		<td width="58" class="titrelog"><a href="images/exit.JPG">
		<img src="http://intra-dev/intrapur/procurement_mrs/ASRaM/images/exit_small.JPG" width="55" height="48" onclick="window.close()" xthumbnail-orig-image="http://intra-dev/intrapur/procurement_mrs/ASRaM/images/exit.JPG" border="2" style="border: 1px solid #FF0000" ></a> 
		</td>
	</tr>
</table>

<table border="0" width="100%">
<tr class="rublog" >
		<td width="309" >
		
		<table border="0" width="100%">
<tr class="rublog" >
		<td width="19%" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		Material <input type="text" name="FIdArticle" size="14" maxlength="14" value="<%=FIdArticle%>"></td>
		<td width="60%" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6"> 
		Material Description :  <input type="text" name="Fdesignation" size="27" maxlength="27" value="<%=FDesignation%>"></td>
<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" >
<input type="submit" value=" Find " name="BFind" style="color: #009900; font-weight: bold"></td>
</tr> 
</table> 
<%
if (Request("Fdesignation") <> "" or FIdArticle <> "") then
%>
<table border="1" width="100%" id="table1" cellspacing="0" cellpadding="0" bordercolordark="#000000">
	<tr class=titrelog  align="center" bgcolor="#DDDDFF">
		<td bgcolor="#DCE9F6">Part Number</td>
		<td width="155" bgcolor="#DCE9F6">IdArticle Description</td>
	</tr>
		<%	

	do until GetArticles.eof
		response.write "<tr class=rublog >" & _
			"<td>"&GetArticles("IdArticle")&"<input type=""checkbox""name=""casacocher""	value="&"'"&GetArticles("IdArticle")&"'"&" onclick=""submit()""></td>" & _
			"<td>&nbsp;<A class=""valider"" href=""#"" onclick=""Submit()"">"&GetArticles("designation")&"</A></td>" & _
		"</tr>" 
	GetArticles.movenext
	loop

	%>
	
</table>
<%end if%>
		
		
	</td>
	<td valign="top">
		<table border="0" width="100%">
			<tr>
				
				<td bgcolor="#DCE9F6" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog">
				Current Material</td>
				<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#DCE9F6">Designation</td>
				<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#DCE9F6">Valid To</td>
				<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#DCE9F6">
				&nbsp;Price</td>
				<td width="43" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#DCE9F6">
				Qty</td>
				<td width="43" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#DCE9F6">
				IdUnit</td>
				<td align=center><input type="submit" value="Save" name="BSave">
				</td>
			</tr>
					<%
		if not GetExistArticle.eof then
		i=1
		do until GetExistArticle.eof
		if GetExistArticle("IdLine")<>IdLine then 'vérif si idLine <> précédente
		%>
		<tr>

			<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
				<input type="text" name="A<%=i%>" value="<%=GetExistArticle("IdArticle")%>" readonly style="background-color: #E3EBEE; text-align:center"></td>
			<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
				<input type="text" readonly name="D<%=i%>" value="<%=trim(GetExistArticle("Designation"))%>" style="background-color: #E3EBEE; text-align:center"></td>
			<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
				<input type="text" readonly name="V<%=i%>" value="<%=trim(GetExistArticle("ValidTo"))%>" style="background-color: #E3EBEE; text-align:center" size="4"></td>
			<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
				<input type="text" name="P<%=i%>" value="<%=GetExistArticle("Art_Price")%>" onkeypress ="javascript:CtrlVal(this)" size="8" style="text-align:center"></td>
			<td width="43" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
			<%if status = "S" then
				response.write	"<input type=""text"" name=" & "Q" & i & " readonly size=""3"" style=""border-style: solid; border-width: 1px; background-color: #E5E5E5; text-align:center"" value=" & GetExistArticle("Art_Qty") & " >"
			else
				response.write	"<input type=""text"" name=" & "Q" & i & " onkeypress =""javascript:CtrlVal(this)"" size=""3"" value=" & GetExistArticle("Art_Qty") & " >"
				end if%>
				<input type="hidden" name="L<%=i%>" value="<%=trim(GetExistArticle("IdLine"))%>"></td>
			<td width="43" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
					
				<select size="1" name="U<%=i%>" >
				<%RsUnit.movefirst%>
				<OPTION Value ="<%=trim(RsUnit("IdUnit"))%>" selected >
				<%=RsUnit("Designation")%></OPTION>
				<%RsUnit.movenext
				do until RsUnit.eof
					if trim(GetExistArticle("IdUnit")) = trim(RsUnit("IdUnit")) then %>
						<OPTION Value ="<%=trim(RsUnit("IdUnit"))%>" selected >
						<%=RsUnit("Designation")%></OPTION>
					<%else%>
						<OPTION Value ="<%=trim(RsUnit("IdUnit"))%>">
						<%=RsUnit("Designation")%></OPTION>
					<%end if
					RsUnit.movenext
				loop%>
				</SELECT>
				
				</td>
			<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" class="rublog" bgcolor="#FFFFFF">
			<a href="Articles_ASRaM.asp?IdBusCase=<%=IdBusCase%>&idLine=<%=trim(GetExistArticle("IdLine"))%>&Status=<%=Status%>&TabInd=<%=TabInd%>&IdCurrency=<%=IdCurrency%>&IdArticle=<%=IdArticle%>&FIdArticle=<%=FIdArticle%>&Fdesignation=<%=Fdesignation%>&Delete='yes'&HIdArticle=<%=trim(GetExistArticle("IdArticle"))%>&ValidTo=<%=trim(GetExistArticle("ValidTo"))%>" >
			<img src="http://intra-dev/intrapur/procurement_mrs/ASRaM/images/delete.JPG" border="0" width="17" height="18">Delete</a>
			
			</td>
			
		</tr>
		<%
		end if
		i=i+1
		IdLine = GetExistArticle("IdLine") 'récup idLine
		GetExistArticle.movenext
		loop
	end if
 %>
		<input type="hidden" name="RsArtCount" value="<%=i%>">
		<input type="hidden" name="HidLine" value="<%=IdLine%>">
		
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