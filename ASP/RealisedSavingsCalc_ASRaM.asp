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

function CtrlVal(){// Procedure de validation
//alert(window.event.keyCode)
if (window.event.keyCode==44 || window.event.keyCode==46)
{
window.event.keyCode = 44;
} else if (window.event.keyCode!=45 && window.event.keyCode!=43 && (window.event.keyCode<48 || window.event.keyCode>57))
{
window.event.keyCode = 0;
}}

function ChpChk(Test, i){
if (Test)
{
document.getElementById('PR_'+i).value=0;
document.getElementById('JB_'+i).disabled=false;
document.getElementById('PR_'+i).disabled=true;
} else {
document.getElementById('JB_'+i).checked=false;
document.getElementById('PR_'+i).disabled=false;
document.getElementById('JB_'+i).disabled=true;
}}

function JBChk(Test, i,BaseYearS,MaxYear){
if (Test)
{for (j=i;j<=MaxYear && document.getElementById('FR_'+(j)).checked; j++)
{document.getElementById('JB_'+(j)).checked=true}
for (j=i;j>BaseYearS && document.getElementById('FR_'+(j)).checked; j--)
{document.getElementById('JB_'+(j)).checked=true}
} else {
for (j=i;j<=MaxYear && document.getElementById('FR_'+(j)).checked; j++)
{document.getElementById('JB_'+(j)).checked=false}
for (j=i;j>BaseYearS && document.getElementById('FR_'+(j)).checked; j--)
{document.getElementById('JB_'+(j)).checked=false}
}}

function ListSelect(Nom,Annee){
if (Annee==BaseYearS)
{
for (i=BaseYearS;i<MaxYear; i++)
{
document.getElementById(Nom+(i+1)).disabled=false
document.getElementById(Nom+(i+1)).selectedIndex=document.getElementById(Nom+Annee).selectedIndex
}
//alert(Test.substr(2, Test.lenght))
}}

</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Realised Savings Calculation</title>
</head>
<form action="RealisedSavingsCalc_ASRaM.asp" name="RealisedSavingsCalc_ASRaM" method="post"  onkeydown ="if (window.event.keyCode == 13) {document.getElementById('BSave').click();}">
<body>
<%

IdBusCase = trim(Request("IdBusCase"))
	
'affiche toute les valeurs envoyée avec la methode post
Response.write "<table border=""1"">"
	For Each chaine_requete In Request.Form
		Response.Write "<tr><td>" & chaine_requete & "</td><td>" & _
			Request.Form(chaine_requete) & "</td></tr>"
	Next
  response.write "</table>///"      


'Fonction Netoie Escal Fmul
function Netoi(Fm)
 	Fm = replace (Fm,"P0;","")
 	Fm = replace (Fm,"P(n-1);","")
 	Fm = replace (Fm,"P(n-2);","")
 	Fm = replace (Fm,"(;","")
	Fm = replace (Fm,");","")
	Fm = replace (Fm,"+;","")
	Fm = replace (Fm,"-;","")
	Fm = replace (Fm,"*;","")
	Fm = replace (Fm,":;","")
end function

MaxYear=2019

set RsArticle =connection.Execute("select * from Adv_Article, Adv_Index_BC_Article where " & _
	"Adv_Article.IdArticle=Adv_Index_BC_Article.IdArticle AND IdBusCase = '" & IdBusCase & _
	"' order by Status, TabInd, IdLine")
set RsEscal =connection.execute("select * from Adv_Escalation where IdBusCase = '" & _
	IdBusCase & "'")
	
do until RsEscal.eof
	Status = trim(RsEscal("Status"))
 	if Status = "B" then 		
		BaseYearB=RsEscal("iYear")
		EscalationFmB = trim(RsEscal("EscalationFm"))
  		AffichEscalationFm=Replace(EscalationFmB,";"," ")
		Netoi EscalationFmB
	elseif Status = "S" then
		BaseYearS=RsEscal("iYear")
		EscalationFmS = trim(RsEscal("EscalationFm"))
 		AffichEscalationFmS=Replace(EscalationFmS,";"," ")
 		Netoi EscalationFmS
	end if
	RsEscal.movenext
loop

'Si PricesRules vide alors reinseigne à vide
Set RsPricesRules =connection.execute("select * from Adv_PricesRules where IdBusCase = '" & _
IdBusCase&"' order by iYear")
if RsPricesRules.eof then
for i = BaseYearS to MaxYear
	ps_adv_MajPricesRules = "ps_adv_MajPricesRules '" &IdBusCase& "','" &i& "','0','','','0',''"
	set AddProc =connection.Execute(ps_adv_MajPricesRules)
next
end if

'Sauver vers BDD ______________________________________________________________
if Request("BSave") = "Save" then 
for n=1 to 2 '1 Tableau par status

	if n=1 then
		tableau = split(EscalationFmB,";")
		Status="B"
	else
		tableau = split(EscalationFmS,";")
		Status="S"
	end if

	IdBusCase=Request("IdBusCase")
	for i = 0 to UBound(tableau, 1)-1
		'Controle si c'est un indice
		if not (isnumeric(tableau(i))) and tableau(i)<>"" and tableau(i) <> "+" and tableau(i) <> "-"  and tableau(i) <> "*"  and tableau(i) <> ":"  and tableau(i) <> "("  and tableau(i) <> ")" then
			for d=BaseYearS to MaxYear 'nombre colonnes 
				IndiceDateMonth=Request("M" & Status& tableau(i)& "-" & BaseYearS)
				IndiceDateYear=Request("Y" & Status & tableau(i)& "-" & BaseYearS)
				IndiceDate = "01/" & IndiceDateMonth & "/" & IndiceDateYear
				iYear=d
				if isnumeric(right(tableau(i),1)) then
					IndicePos = right(tableau(i),1)
					IdIndice = left(tableau(i),len(tableau(i))-1)
				else
					IndicePos = 9999
					IdIndice = tableau(i)
				end if
				ps_adv_MajIndices_Month = "ps_adv_MajIndices_Month '" &IdBusCase& "','" &Status& "','" & _
		 		IdIndice& "','" &IndicePos& "','" &iYear& "','" &formatDate(IndiceDate)& "',''"
				Response.write ps_adv_MajIndices_Month
				set AddProc =connection.Execute(ps_adv_MajIndices_Month)

			next
		end if
	next
next

	Set RsPricesRules =connection.execute("select * from Adv_PricesRules where IdBusCase = '" & _
	IdBusCase&"' order by iYear")
	for i = BaseYearS to MaxYear
		PriceReduction=replace (Request("PR_" & i),",",".")
		Freeze = Request("Fr_" & i)
		JumpBack = Request("JB_" & i)
		AdvanceSavings = replace (Request("AS_" & i),",",".")
	
		ps_adv_MajPricesRules = "ps_adv_MajPricesRules '" &IdBusCase& "','" &i& "','" & _
			PriceReduction& "','" &Freeze& "','" &JumpBack& "','" &AdvanceSavings& "',''"
		set AddProc =connection.Execute(ps_adv_MajPricesRules)
	next

end if
%>
<input type="hidden" name="IdBusCase" value = "<%=IdBusCase%>"  readonly size="19">

	<table border="2" width="100%" style="border: 3px ridge #CCCCFF">
		<tr>
			<td width="25%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
			<table width="740" height="53">
			<tr><!-- Ligne suivante ---------------------------------->
				<td width="186" height="24" style="border-style: ridge; border-width: 2px; padding: 0" bordercolordark="#C0C0C0"> 
				Base line Escalation Formula : </td>
				<td width="540" height="24" style="border-style: ridge; border-width: 2px; padding: 0" bordercolordark="#C0C0C0">
				<%=AffichEscalationFm%>
				</td>
			</tr>	
			<tr>
				<td width="186" height="23" style="border-style: ridge; border-width: 2px; padding: 0" bordercolordark="#C0C0C0"> 
				Secured Escalation Formula : </td>
				<td width="540" height="23" style="border-style: ridge; border-width: 2px; padding: 0" bordercolordark="#C0C0C0">
				<%=AffichEscalationFmS%>
				</td>
			</tr>	
			</table>
			</td>
		</tr>	
		<tr><!-- Ligne suivante ---------------------------------->

		<%for n=1 to 2 'nombre Tableaux%>
			<!-- Cree un tableau -->
			<table border="2" width="100%" id="tablo" style="border: 3px ridge #CCCCFF">
			<p>
				<td nowrap width="140"> 
			<%				
			if n=1 then
			tableau = split(EscalationFmB,";")
			Status="B"
			Response.write "Base line"
			else
			tableau = split(EscalationFmS,";")
			Status="S"
			Response.write "Secured"
			end if
			%>
</td>
<%			for i = BaseYearS to MaxYear%>
				<td width ="36" style="background-color: #B0CFD0"> <!-- Cree une ligne -->
					<input type="text" name="HC_<%=n & "_" & i-BaseYearS%>"  size="13" readonly
					value ="<%=i%>" style="background-color: #F0F0F0; padding-left:3px">
				</td>
			<%next
			'Efface les Indices en double
			for i = 0 to UBound(tableau, 1)
				'Efface P0, P1, P2
				if tableau(i)="P0" or tableau(i)="P(n-1)" or tableau(i)="P(n-2)" then tableau(i)=""
				for j = UBound(tableau, 1) to 0 step -1
					if i<>j and tableau(i)=tableau(j) then
						tableau(j)=""
					end if
				next
			next
			
			for i = 0 to UBound(tableau, 1)-1
			if not (isnumeric(tableau(i))) and tableau(i)<>"" then
					
				%>
				<tr> <!-- Cree une colonne -->
			<%for d=BaseYearS-1 to MaxYear 'nombre colonnes
				if tableau(i) <> "+" and tableau(i) <> "-"  and tableau(i) <> "*"  and tableau(i) <> ":"  and tableau(i) <> "("  and tableau(i) <> ")" then
				'Controle si c'est un indice
					if isnumeric(right(tableau(i),1)) then
						IndicePos = right(tableau(i),1)
						IdIndice = left(tableau(i),len(tableau(i))-1)
					else
						IndicePos = 9999
						IdIndice = tableau(i)
					end if
					
				set RsIndices =connection.execute("select * from Adv_Indices_month where IdBusCase = '" & _
				IdBusCase&"' and Status = '"&Status&"' and IdIndice = '"&IdIndice&"' and IndicePos = "&IndicePos& _
				" and iYear = "& (d) &" order by Status")
				%>
				<td nowrap style="background-color: #B0CFD0"> <!-- Cree une ligne -->
					<%If d=BaseYearS-1 then '1ere colonne uniquement %>
						<input type="text" name="HL_<%=n & "_" & i%>"  size="19" 
						value ="<%=tableau(i)%>" style="background-color: #F0F0F0" readonly="true">
					
					<%ElseIf not RsIndices.eof then 'si date existante dans la BDD
						Indat=RsIndices("IndiceDate")
			%>
						<select size="1" name="M<%=Status&tableau(i) & "-" & d%>"  
						 onchange ="ListSelect('M<%=Status&tableau(i)%>-',<%=d%>)">
							<%for j = 1 to 12
								if j=month(Indat) then
									if j<10 then%>
										<OPTION Value ="0<%=j%>" selected >
										0<%=j%></OPTION>
									<%Else%>
										<OPTION Value ="<%=j%>" selected >
										<%=j%></OPTION>
									<%end if
								Elseif j=1 then%>
									<OPTION Value ="01" selected >
									01</OPTION>
								<%Elseif j<10 then%>
									<OPTION Value ="0<%=j%>">
									0<%=j%></OPTION>
								<%Else%>
									<OPTION Value ="<%=j%>">
									<%=j%></OPTION>
								<%end if
							next%>
						</SELECT>

						<select size="1" name="Y<%=Status&tableau(i) & "-" & d%>">
							<%
							for j = 1 to MaxYear-BaseYearS+2
								if j=year(Indat)-BaseYearS then%>
									<OPTION Value ="<%=BaseYearS+j%>" selected >
									<%=BaseYearS+j%></OPTION>
								<%else%>
									<OPTION Value ="<%=BaseYearS + j%>">
									<%=BaseYearS + j%></OPTION>
								<%end if
							next%>
						</SELECT>

					<%Else 'pas existante dans BDD%>
						<select size="1" name="M<%=Status&tableau(i) & "-" & d%>"   
						 onchange ="ListSelect('M<%=Status&tableau(i)%>-',<%=d%>)">
							<%for j = 1 to 12
								if j=1 then%>
									<OPTION Value ="01" selected >
									01</OPTION>
								<%Elseif j<10 then%>
									<OPTION Value ="0<%=j%>">
									0<%=j%></OPTION>
								<%Else%>
									<OPTION Value ="<%=j%>">
									<%=j%></OPTION>
								<%end if
							next%>
							
						</SELECT>
						<select size="1" name="Y<%=Status&tableau(i) & "-" & d%>">
							<%for j = 1 to MaxYear-BaseYearS+2
								if j=d-BaseYearS then 'année en cours
									if IndicePos = 0 then%>
										<OPTION Value ="<%=BaseYearS%>" selected >
										<%=BaseYearS%></OPTION>
									<%ElseIf IndicePos = 9999 then%>
										<OPTION Value ="<%=BaseYearS + j%>" selected >
										<%=d%></OPTION>
									<%ElseIf IndicePos = 1 then%>
										<OPTION Value ="<%=BaseYearS + j%>" selected >
										<%=BaseYearS + j%></OPTION>
									<%ElseIf IndicePos = 2 then%>
										<OPTION Value ="<%=BaseYearS-4 + j%>" selected >
										<%=BaseYearS-4 + d%></OPTION>
									<%Else%>
										<OPTION Value ="<%=BaseYearS + j%>" selected >
										<%=BaseYearS + j%></OPTION>
									<%end if
								else%>
									<OPTION Value ="<%=BaseYearS + j%>">
									<%=BaseYearS + j%></OPTION>
								<%end if
							next%>
						</SELECT>					
										
					<%end if%>
					</td>
				<%end if
				next%>
				</tr>
			<%end if
			next
			%>
			</table>
		<%next
		Set RsPricesRules =connection.execute("select * from Adv_PricesRules where IdBusCase = '" & _
				IdBusCase&"' order by iYear")
		%>
	<table border="2" width="100%" style="border: 3px ridge #CCCCFF">
		<tr>
			<td nowrap width="140">
				Price Variation (%)</td>
			<%do until RsPricesRules.eof
			Test=trim(RsPricesRules("iYear"))
			for i = BaseYearS to MaxYear
				if cstr(i) = Test then 
				if RsPricesRules("Freeze") = "x" then %>
					<td style="background-color: #B0CFD0" nowrap width="60" align="right"> <!-- Cree une ligne -->
						<input type="text" disabled name="PR_<%=i%>" value ="<%=RsPricesRules("PriceReduction")%>" 
						onkeypress ="javascript:CtrlVal()" size="13"  style="background-color: #F0F0F0">
					</td>
				<%else%>
					<td style="background-color: #B0CFD0" nowrap width="60" align="right"> <!-- Cree une ligne -->
						<input type="text" name="PR_<%=i%>" value ="<%=RsPricesRules("PriceReduction")%>" 
						onkeypress ="javascript:CtrlVal()" size="13"  style="background-color: #F0F0F0">
					</td>
				<%end if
				end if
			next
			RsPricesRules.movenext
			loop%>
		</tr>
	</table>
	<table border="2" width="100%" style="border: 3px ridge #CCCCFF">
		<tr>
			<td nowrap width="140">
			Freeze</td>
			<td style="background-color: #B0CFD0" nowrap width="105" align="center"></td>
			<%RsPricesRules.movefirst
			do until RsPricesRules.eof
			Test=trim(RsPricesRules("iYear"))
			for i = BaseYearS+1 to MaxYear+1
				if cstr(i) = Test then
				if RsPricesRules("Freeze") = "x" then %>
					<td style="background-color: #B0CFD0" nowrap width="105" align="center"> <!-- Cree une ligne -->
						<input type="checkbox" checked name="Fr_<%=i%>" value="x" 
						 onclick="ChpChk(this.checked,<%=i%>)">
					</td>
				<%else%>
					<td style="background-color: #B0CFD0" nowrap width="105" align="center"> <!-- Cree une ligne -->
						<input type="checkbox" name="Fr_<%=i%>" value="x"
						onclick="ChpChk(this.checked, <%=i%>)">
					</td>
				<%end if
				end if
			next
			RsPricesRules.movenext
			loop%>
		</tr>
		<tr>
			<td nowrap width="140">
				With Jump Back</td>
			<td style="background-color: #B0CFD0" nowrap width="60" align="right"></td>
			<%RsPricesRules.movefirst
			do until RsPricesRules.eof
			Test=trim(RsPricesRules("iYear"))
			for i = BaseYearS+1 to MaxYear+1
				if cstr(i) = Test then
				if RsPricesRules("JumpBack") = "x" then %>
					<td style="background-color: #B0CFD0" nowrap width="60" align="right"> <!-- Cree une ligne -->
						<input type="checkbox" checked name="JB_<%=i%>" value="x"
						onclick="JBChk(this.checked,<%=i%>,<%=BaseYearS%>,<%=MaxYear%>)">
					</td>
				<%elseif RsPricesRules("Freeze") = "x" then %>
					<td style="background-color: #B0CFD0" nowrap width="60" align="right"> <!-- Cree une ligne -->
						<input type="checkbox" name="JB_<%=i%>" value="x"
						onclick="JBChk(this.checked,<%=i%>,<%=BaseYearS%>,<%=MaxYear%>)">
					</td>	
				<%else%>
					<td style="background-color: #B0CFD0" nowrap width="60" align="right"> <!-- Cree une ligne -->
						<input type="checkbox" disabled name="JB_<%=i%>" value="x"
						onclick="JBChk(this.checked,<%=i%>,<%=BaseYearS%>,<%=MaxYear%>)">
					</td>
				<%end if
				end if
			next
			RsPricesRules.movenext
			loop%>
		</tr>
	</table>
	<table border="2" width="100%" style="border: 3px ridge #CCCCFF">
		<tr>
			<td nowrap width="140">
				ADVANCE Business Case Secured Savings</td>
			<%RsPricesRules.movefirst
			do until RsPricesRules.eof
			Test=trim(RsPricesRules("iYear"))
			for i = BaseYearS to MaxYear
				if cstr(i) = Test then %>
					<td style="background-color: #B0CFD0" nowrap width="36" align="right"> <!-- Cree une ligne -->
						<input type="text" name="AS_<%=i%>"  size="13" onkeypress ="javascript:CtrlVal()"
						value ="<%=RsPricesRules("AdvanceSavings")%>" style="background-color: #F0F0F0">
					</td>
				<%end if
			next
			RsPricesRules.movenext
			loop%>
		</tr>
	</table>
	<table border="2" width="156" style="border: 3px ridge #CCCCFF">
	<tr>
		<td align="center">
		<input type="reset" value="Rétablir" name="B2"><input type="submit" value="Save" name="BSave" style="color: #FF0000">
		</td>
	</tr>
	</table>

</body>
<script language="JavaScript">
<!--#INCLUDE FILE="connection/deconnection.asp"-->
</script>
</form>
</html>