<html>
<script language="javascript">
<!--#INCLUDE FILE="connection/connection.asp"-->

function CtrlVal(Nom){// Procedure de validation
//alert(window.event.keyCode)
if (window.event.keyCode == 13) {FillFmul('tVal')}
if (window.event.keyCode==46 || window.event.keyCode==44)
{
window.event.keyCode = 44;
} else if (window.event.keyCode<48 || window.event.keyCode>57)
{
window.event.keyCode = 0;
}}

function DelFmul(){// Procedure d'effacement
var var2=document.getElementById('hFmul').value; var var1=''
document.getElementById('hFmul').value = var2.substring(0,1+var2.lastIndexOf(";",var2.length-2))
document.getElementById('EscalationFm').value=''
ShowFmul();
}

function FillFmul(Nom){// Procedure de Remplissage
	if (document.getElementById(Nom).value=="/")
	{
	document.getElementById('hFmul').value = document.getElementById('hFmul').value + ":;";
	} else {
	document.getElementById('hFmul').value = document.getElementById('hFmul').value + document.getElementById(Nom).value + ";";
	}
	ShowFmul();
}

function ShowFmul(){// Procedure de Lecture
	var var2=''; var var1 = ''
	document.getElementById('EscalationFm').value=''
	var2= document.getElementById('hFmul').value
	while (var1.length<var2.length) {
		var1=document.getElementById('EscalationFm').value
		if (var1.length> 0)
			{// si EscalationFm renseigné
			document.getElementById('EscalationFm').value = var1 + var2.substring(1+var2.lastIndexOf(";",var1.length),var2.indexOf(";",var1.length)) + " "
		}else {// si EscalationFm vide
			document.getElementById('EscalationFm').value = var2.substring(0,var2.indexOf(";",1)) + " "
		}
		var1=document.getElementById('EscalationFm').value
	}
}

// -----------------------------------------------------------
</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Escalation Formula</title>
<%' _________ FONCTIONS __________
%>
</head>

<%
'affiche toute les valeurs envoyée avec la methode get 
'  For Each chaine_requete In Request.QueryString
' 		Response.Write "<tr>;<td> - " & chaine_requete & "</td> = <td>" & Request.QueryString(chaine_requete) & "</td></tr>" 
'  Next

	set RsIndices =connection.execute("select * from Adv_Indice order by IdIndice")
	IdBusCase= trim(Request.querystring("IdBusCase"))
	Status = trim(Request.querystring("Status"))
	Bsave  = trim(Request.querystring("Bsave"))
	hFmul = trim(Request.querystring("hFmul"))
	iYear = trim(Request.querystring("iYear"))
	
	
	if Bsave <> "" then 'Sauvegarde de la Formule
		ps_Adv_AddEscalation = "ps_Adv_AddEscalation '" & IdBusCase & "','" & Status & "','" & hFmul & "','" & iYear & "'"
			set Maj=connection.Execute(ps_Adv_AddEscalation)
		%>		
		<script language="javascript">
		window.opener.location.reload()
		window.close()			
		</script>	
	<%end if

		TesSQL ="select * from Adv_Escalation where IdBusCase = '" & IdBusCase & "' and Status = '" & Status & "' order by IdBusCase"
		set RsEscal =connection.execute(TesSQL)
	 	if not (RsEscal.eof and RsEscal.bof) then ' Fmul existe
	 		RsEscal.movefirst
			hFmul = trim(RsEscal("EscalationFm"))		
		end if

%>
<body >

<!-- Cree un form: focus sur controle et valide si on presse enter -->

<form name="Escal" method="get">

<input type="hidden" name="hFmul" size="10" value="<%=hFmul%>"><p>

<table border="1" width="100%" id="table8" >
	<tr>
		<td width="251">Busness Case N° 
		<input type="text" name="IdBusCase" value = "<%=IdBusCase%>"  readonly size="19"></td>
		<td>Status<input type="text" name="Status" value = "<%=Status%>"  readonly size="19"></td>
		<td >
		<img style="border-style: solid; border-width: 0px" border="0" src="images/Exit_mini.JPG" width="25" height="24" align="right"
		onclick=" window.close()"></td>
		
	</tr>
</table>

<table border="1" width="100%">
	<tr>
		<td width="95%">
		<input type="text" name="EscalationFm" value = "<%=trim(Replace(hFmul,";"," "))%>" readonly size="120%">
		</td>
	</tr>
</table>

<table border="1" width="100%" align="left" >
	<tr>
		<td height="20" align="right" valign="top">
		<table border="1" align="right" id="table12">
			<tr>
				<td width="42" height="27">
				<input type="button" value="ç" name="BTdel" onclick="Javascript:DelFmul()" 
				style="font-family: Wingdings; font-size: 12px; color: #FF0000; padding-left:0px; padding-right:1px; padding-top:2px; padding-bottom:1px; width:40; font-weight:bold">
				</td>
			</tr>
			<tr>
				<td height="20" width="42">
				<input type="button" value="P0" name="P0" style="background-color: #CCCCFF; font-size:8pt" 
				onclick="Javascript:FillFmul('P0')"></td>
			</tr>
			<tr>
				<td width="42" height="28"><input type="button" value="P(n-1)" name="Pn1" style="background-color: #CCCCFF; font-size:8pt" 
				onclick="Javascript:FillFmul('Pn1')"></td>
			</tr>
			<tr>
			<td width="42">
			<input type="button" value="P(n-2)" name="Pn2" style="background-color: #CCCCFF; font-size:8pt" 
				onclick="Javascript:FillFmul('Pn2')"></td>
		</tr>
	</table>

		</td>
		<td align="right" valign="top">

		<table border="1" align="right" id="table11" width="63">
			<tr>
				<td height="27" colspan ="2">
				
	<table border="1" width="100%" id="table13" bordercolor="#008000">
		<tr>
			<td>Valeur manuelle:</td>
		<tr>
			<td>
			<input type="text" name="tVal" size="14" onkeypress ="javascript:CtrlVal('tVal')"></td>
		</tr>
			<td><input type="button" value="Add" name="Manu" style="background-color: #CCCCFF; font-family:Microsoft Sans Serif; font-size:8pt; font-weight:bold; padding-left:1px; padding-right:1px; padding-top:0; padding-bottom:0" 
			onclick="Javascript:FillFmul('tVal')"></td>
		</tr>

	</table>
				
				
				
				</td>
			</tr>
			<tr>
				<td width="24"><input type="button" value="+" name="plus" style="background-color: #FFCC99; color:#0000FF; font-weight:bold; font-size:8pt; padding-left:1px; padding-right:1px; padding-top:0px; padding-bottom:0px" 
				onclick="Javascript:FillFmul('plus')"></td>
				<td width="23"><input type="button" value="-" name="moins"  style="background-color: #FFCC99; color:#0000FF; font-size:8pt; font-family:Symbol; font-weight:bold; padding-left:1px; padding-right:1px; padding-top:0; padding-bottom:0" 
				onclick="Javascript:FillFmul('moins')"></td>
			</tr>
			<tr>
			<td width="24" height="28"><input type="button" value="*" name="fois" style="background-color: #FFCC99; color:#0000FF; font-family:Symbol; font-size:8pt; font-weight:bold; padding-left:1px; padding-right:1px; padding-top:0; padding-bottom:0" 
			onclick="Javascript:FillFmul('fois')"></td>
			<td height="28" width="23">
			<input type="button" value="/" name="div" 
			style="background-color: #FFCC99; color:#0000FF; font-size:8pt; font-weight:bold; padding-left:2px; padding-right:2px; padding-top:0; padding-bottom:1px; font-family:Symbol" onclick="Javascript:FillFmul('div')"></td>
			</tr>
			<tr>
			<td width="24">
			<input type="button" value="(" name="OpenBracket" style="background-color: #FFCC99; color:#0000FF; font-size:8pt; font-weight:bold; font-family:Symbol; padding-bottom:1px; padding-left:1px; padding-right:1px; padding-top:0" 
			onclick="Javascript:FillFmul('OpenBracket')"></td>
			<td width="23">
			<input type="button" value=")" name="CloseBracket" style="background-color: #FFCC99; color:#0000FF; font-family:Symbol; font-size:8pt; font-weight:bold; padding-top:0; padding-bottom:1px; padding-left:1px; padding-right:1px" 
			onclick="Javascript:FillFmul('CloseBracket')"></td>
		</tr>

	</table>

		</td>
		<td rowspan="4" width="90%" align="left" valign="top">
		<table border="1" width="100%" id="table6">
			<tr>
			<td  height="26" align="left" colspan="2" title="Year where the calculs start">Base Line Year
				<select size="1" name="iYear" title="Year where the calculs start" >
				<OPTION></OPTION>
					<% if not RsEscal.eof then
						'il existe une formule
						
						For i = 1999 to Year(Date())+1
							if trim(RsEscal("iYear")) = trim(i) then
							%>
							<OPTION Value ="<%=i%>" selected ><%=i%></OPTION>
							<%
							else 
							%>
							<OPTION Value ="<%=i%>"><%=i%></OPTION>
							<%
							end if
						next
					else
						%>
						<OPTION Value ="1999" selected >1999</OPTION>
						<%
						For i = 2000 to Year(Date())+1
							%>
							<OPTION Value ="<%=i%>"><%=i%></OPTION>
							<%	
						next
				end if
				%>
				</SELECT>
			
			</td>
			</tr> 
			<tr>
			<td  height="26" align="left">Base Line Year Indice</td>
			<td  height="26" align="left">Current Year</td>
			<td align="left">Year - 1</td>
			<td align="left">Year - 2</td>
			</tr>
		<%RsIndices.movefirst
		do until RsIndices.eof%>
			<tr>
			<td width="100" align="left">
			<input type="button" value="<%=trim(RsIndices("IdIndice"))%>0" name="<%=trim(RsIndices("IdIndice"))%>0" style="background-color: #CCCCFF; font-size:8pt; font-family:Microsoft Sans Serif; font-weight:bold" 
			onclick="Javascript:FillFmul('<%=trim(RsIndices("IdIndice"))%>0')"></td>
			<td width="50" align="left">
			<input type="button" value="<%=trim(RsIndices("IdIndice"))%>" name="<%=trim(RsIndices("IdIndice"))%>" style="background-color: #CCCCFF; font-size:8pt; font-family:Microsoft Sans Serif; font-weight:bold" 
			onclick="Javascript:FillFmul('<%=trim(RsIndices("IdIndice"))%>')"></td>
			<td width="50" align="left">
			<input type="button" value="<%=trim(RsIndices("IdIndice"))%>1" name="<%=trim(RsIndices("IdIndice"))%>1" style="background-color: #CCCCFF; font-size:8pt; font-family:Microsoft Sans Serif; font-weight:bold" 
			onclick="Javascript:FillFmul('<%=trim(RsIndices("IdIndice"))%>1')"></td>
			<td width="50" align="left">
			<input type="button" value="<%=trim(RsIndices("IdIndice"))%>2" name="<%=trim(RsIndices("IdIndice"))%>2" style="background-color: #CCCCFF; font-size:8pt; font-family:Microsoft Sans Serif; font-weight:bold" 
			onclick="Javascript:FillFmul('<%=trim(RsIndices("IdIndice"))%>2')"></td>
			<td><%=trim(RsIndices("Designation"))%></td>
			</tr>
			<%RsIndices.movenext
		loop%>
		</table>
		</td>
		</tr>
		
	<tr>
		</tr>
		
	<tr height="40">
		<td height="80%" align="right" valign="top" colspan="2">

	<table border="0"  width="100%" style="border-style: outset; border-width: 1px; background-color: #FFCC99" id="table14">
		<tr>
			<td>
			<input type="reset" value="Clear" name="B2"style="color:#FF5050; font-weight:bold"></td>
			<td><input type="submit" value="Save" name="BSave" style="color:#0000FF"></td>
		</tr>
	</table>
		</td>
		</tr>
		
	<tr >
		</tr>
		
		</table>
	</table>


<%if request(BSave2) = "Save2" then 'Calculs ________________________________
'Remplit le Tableau
	tableau = split(hFmul,";")
	i=0
	while tableau(i) <> ""
	i=i+1
	wend

	response.write "Calc" & hFmul %></p>
<p>
	&nbsp;</p><%
	
'Calcule
	LngFmul=i-1
	VarBool = true
	while VarBool = true
		VarBool = false
		For i = 0 to LngFmul
			if tableau(i) = "(" then
				ValOpen=i
				VarBool = true
			end if
		next
		if VarBool = true then
			i=ValOpen
			while tableau(i) <> ")" and i < LngFmul%>
	
		<table>
			<tr>
	
<%				n3 = i
				while n3 < LngFmul and tableau(n3) <> ")"
				if tableau(n3) = "*" or tableau(n3) = ":" then
					if IsNumeric(tableau (n3-1)) and IsNumeric(tableau (n3+1)) then _
					Calcule tableau(n3-1), tableau(n3+1), tableau(n3)
					RefreshTab LngFmul
				else
					n3=n3+1
				end if
				wend				
				if IsNumeric(tableau (i)) then
					Calcule tableau(i), tableau(i+2), tableau(i+1)
				end if
				i=i+1
			wend
			tableau(i) = ""
			tableau(ValOpen) = ""
			RefreshTab LngFmul
		end if
		
for n2 = 0 to LngFmul
	response.write i & "N"& n2 & " : " & tableau(n2)%></p><%
next
%>
		</tr>
	</table>
	
<%	
	wend

	n3=0
	while n3 < LngFmul and tableau(n3) <> ")"
		if tableau(n3) = "*" or tableau(n3) = ":" then
			if IsNumeric(tableau (n3-1)) and IsNumeric(tableau (n3+1)) then _
			Calcule tableau(n3-1), tableau(n3+1), tableau(n3)
			RefreshTab LngFmul
		else
				n3=n3+1
		end if
	wend				
i=0
	while tableau(i+2) <> "" and i < LngFmul
		if IsNumeric(tableau (i)) then
			Calcule tableau(i), tableau(i+2), tableau(i+1)
			RefreshTab LngFmul
		else
			i=i+1
		end if
	wend
for n2 = 0 to LngFmul
	response.write i & "N"& n2 & " : " & tableau(n2)%></p><%
next

end if%>


<%' _________ FONCTIONS __________
function Calcule(Val1 , Val2, Signe) ' ____________________________________________________
select case Signe
case "*"
Val1 = Val2 *Val1
Val2 = ""
Signe=""
case ":"
Val1 = Val1 / Val2
Val2 = ""
Signe=""
case "+"
Val1 = CDbl(Val1) + CDbl(Val2)
Val2 = ""
Signe=""
case "-"
Val1 = CDbl(Val1) - CDbl(Val2)
Val2 = ""
Signe=""
end select
End function

function RefreshTab(Maxx)
for n = 0 to Maxx
	if tableau(n) = "" then
		n1=0
		while tableau(n1+n) = "" and n1+n< Maxx
			n1 = n1+1
		wend
		if tableau(n1+n) <> "" then
			tableau(n) = tableau(n1+n)
			tableau(n1+n) = ""
		end if
		
	end if
	if n1+n = Maxx then exit for
next
end function
%>

</form>

</body>

<script language="JavaScript">
<!--#INCLUDE FILE="connection/deconnection.asp"-->
</script>

</html>