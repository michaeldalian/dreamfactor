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

function CtrlVal(Nom){// Procedure de validation
//alert(window.event.keyCode)
if (window.event.keyCode==44 || window.event.keyCode==46)
{
window.event.keyCode = 44;
} else if (window.event.keyCode<48 || window.event.keyCode>57)
{
window.event.keyCode = 0;
}}

</script>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>Reporting</title>
</head>
<body>
<form action="Report_ASRaM_Detail.asp" name="Report_ASRaM_Detail.asp" method="post">

<%' _________ FONCTIONS __________
function Calcule(Val1 , Val2, Signe) ' ______________
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
End function ' ______________________________________

function RefreshTab(Maxx) ' __________________________
for NumTemp = 0 to Maxx
	if Tableau(NumTemp) = "" then
		NumTemp1=0
		while Tableau(NumTemp1+NumTemp) = "" and NumTemp1+NumTemp< Maxx
			NumTemp1 = NumTemp1+1
		wend
		if Tableau(NumTemp1+NumTemp) <> "" then
			Tableau(NumTemp) = Tableau(NumTemp1+NumTemp)
			Tableau(NumTemp1+NumTemp) = ""
		end if
	end if
	if NumTemp1+NumTemp = Maxx then exit for
next
end function ' ______________________________________

function CalculTab(Tablo) ' __________________________
		mcpt=0
		LngFmul=UBound(Tablo, 1)-1
	
		VarBool = true
		Probleme=0
		while VarBool = true AND Probleme < 500
			VarBool = false
			For mcpt1 = 0 to LngFmul
				if Tablo(mcpt1) = "(" then
					ValOpen=mcpt1
					VarBool = true
				end if
			next
			if VarBool = true then
				mcpt=ValOpen
				while Tablo(mcpt) <> ")" and mcpt < LngFmul
					ncpt = mcpt
					Probleme2=0
					while ncpt < LngFmul and Tablo(ncpt) <> ")" AND Probleme2 > 500 
					if Tablo(ncpt) = "*" or Tablo(ncpt) = ":" then
						if IsNumeric(Tablo (ncpt-1)) and IsNumeric(Tablo (ncpt+1)) then _
						Calcule Tablo(ncpt-1), Tablo(ncpt+1), Tablo(ncpt)
						RefreshTab LngFmul
					else
						ncpt=ncpt+1
					end if
					Probleme2 = 1 + Probleme2
					wend
					Probleme2=0			
					if IsNumeric(Tablo (mcpt)) then
						Calcule Tablo(mcpt), Tablo(mcpt+2), Tablo(mcpt+1)
					end if
					mcpt=mcpt+1
				wend
				Tablo(mcpt) = ""
				Tablo(ValOpen) = ""
				RefreshTab LngFmul
			end if
			Probleme = 1 + Probleme
		wend
		
		ncpt=0
		Probleme=0
		while ncpt < LngFmul and Tablo(ncpt) <> ")" AND Probleme < 500
			if Tablo(ncpt) = "*" or Tablo(ncpt) = ":" then
				if IsNumeric(Tablo (ncpt-1)) and IsNumeric(Tablo (ncpt+1)) then _
				Calcule Tablo(ncpt-1), Tablo(ncpt+1), Tablo(ncpt)
				RefreshTab LngFmul
			else
				ncpt=ncpt+1
			end if
		wend				
		Probleme = 0
		
		mcpt=0

		Probleme=0
		while Tablo(mcpt+2) <> "" and mcpt < LngFmul AND Probleme < 500
			if IsNumeric(Tablo (mcpt)) then
				Calcule Tablo(mcpt), Tablo(mcpt+2), Tablo(mcpt+1)
				RefreshTab LngFmul
			else
				mcpt=mcpt+1
			end if
		wend
end function ' Fin des Fonctions __-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-____________________

'affiche toute les valeurs envoyée avec la methode post
'response.write "<table>"
'      For Each chaine_requete In Request.Form
'       Response.Write "<tr><td>" & chaine_requete & "</td><td>" _
'                   & Request.Form(chaine_requete) & "</td></tr>"
'     Next
'  response.write "</table>"      

ReDim P(100,0), Pb(100,0), DvScQty(100,0), PxSAP(100,0), AmountSAP(100,0)
MaxYear=2009
IdBusCase = trim(Request("IdBusCase"))

Set RsPricesRules =connection.execute("select * from Adv_PricesRules where IdBusCase = '" & _
	IdBusCase&"' order by iYear")
'set RsCurrencyRate =connection.execute("select * from Adv_Currency_Rate order by iYear")
Set RsArticles = connection.execute("select * from Adv_Index_BC_Article where IdBusCase = '" & _
	IdBusCase&"' order by TabInd")

set RsEscal =connection.execute("select * from Adv_Escalation where IdBusCase = '" & _
	IdBusCase & "'")
	
do until RsEscal.eof
	Status = trim(RsEscal("Status"))
 	if Status = "B" then 		
		EscalationFmB = trim(RsEscal("EscalationFm"))
		BaseYearB=RsEscal("iYear")
	elseif Status = "S" then
		EscalationFmS = trim(RsEscal("EscalationFm"))
		BaseYearS=RsEscal("iYear")
	end if
	RsEscal.movenext
loop

RsArticles.movefirst
TabInd=RsArticles("TabInd")
do Until RsArticles.eof 
%>
<!-- Cree un Tableau -->
<table border="2" width="100%" style="border: 3px ridge #CCCCFF">
	<input type="hidden" name="IdBusCase" value = "<%=IdBusCase%>"  readonly size="19">
<%	
	do Until RsArticles("TabInd")<>TabInd
		'Sommes prix
		if RsArticles("ValidTo")="9999" AND RsArticles("Status")="B" then	P0b=P0b+RsArticles("Art_Price")
		if RsArticles("ValidTo")="9999" AND RsArticles("Status")="S" then	P0s=P0s+RsArticles("Art_Price")
	
		TabInd=RsArticles("TabInd")
		RsArticles.movenext
		if RsArticles.eof then exit do
	loop

	ReDim Preserve P(100,TabInd)
	ReDim Preserve Pb(100,TabInd)
	ReDim Preserve DvScQty(100,TabInd)
	ReDim Preserve PxSAP(100,TabInd)
	ReDim Preserve AmountSAP(100,TabInd)

	set RsGpeVar =connection.Execute("SELECT distinct IdVariant FROM dbo.Adv_Index_BC_Variant WHERE IdBusCase = '"& IdBusCase &"' AND TabInd = "& TabInd &" AND Status = 'S' ORDER BY IdVariant")
	TabGpeVar=RsGpeVar.GetRows() ' récup Rs dans Tab

	response.write	"<tr>" & _
	"<td height=""5"" align=""center"" colspan=""5"">Variant Group Number "&TabInd&" : "
	z1=0
	'do Until RsGpeVar.eof
	'	response.write	RsGpeVar("IdVariant") : </td></tr>"
	'RsGpeVar.movenext
	do while z1 < UBound(TabGpeVar, 2) and z1<500
		response.write	TabGpeVar(0,z1) & " / "
		z1=z1+1
	loop
		response.write	TabGpeVar(0,z1)
		response.write	"</td></tr>"
		set RsVariant =connection.Execute("SELECT YEAR(dbo.Adv_DelivSchedule.iDate) AS CurrentYear, dbo.Adv_Index_BC_Variant.IdVariant, SUM(dbo.Adv_DelivSchedule.Qty) AS SommeDeQty FROM Adv_DelivSchedule, Adv_Index_BC_Variant WHERE Adv_DelivSchedule.IdVariant = Adv_Index_BC_Variant.IdVariant AND IdBusCase = '"& IdBusCase &"' AND TabInd = "& TabInd &" AND Status = 'S' GROUP BY YEAR(dbo.Adv_DelivSchedule.iDate), dbo.Adv_Index_BC_Variant.IdVariant ORDER BY YEAR(dbo.Adv_DelivSchedule.iDate)")

	for n=1 to 10 'nombre de Lignes
		
		Select case n
		case 1
		Name="Contract Base Year "
		case 2
		Name="Theorical Price Before action"
		case 3
		Name="Theorical Price After action"
		case 4
		Name="Qty of shipset from variant delivery schedule"
		case 5
		Name="Theorical saving based on variant BC calculation"
		case 6
		Name="SAP Negociated Price"
		case 7
		Name="Negociated price monitoring against theorical price"
		case 8
		Name="Theorical saving based on SAP Negociated Price"
		case 9
		Name="Theorical saving based on SAP Negociated Price"
		case 10
		Name="Delta Good receipt price per shipset and theorical price"
		end select
		
		Select case n
	case 1
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"" bordercolorlight=""#800080"" bordercolordark=""#000080"" style=""background-color: #E9E1FF"">" & _
			Name&" :<br>"&BaseYearS &"</td><td nowrap width=""140"" align=""center"" bordercolorlight=""#800080"" bordercolordark=""#000080"" style=""background-color: #E9E1FF"">" & _
			"Contract Base Price<br>P<SUB>0</SUB> :<br>"&P0b & "</td>"
		
		for i = BaseYearS to MaxYear
			response.write	"<td nowrap width =""100"" style=""background-color: #B0CFD0"" align=""center"">"& _
			i & "</td>"
		next
		
	case 2 'Theorical Price Before action _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">A</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
		
			response.write "<td title = ""Result"">"
			Pb(0,TabInd)=P0b
			response.write P0b
			response.write "</td>"
					
		' ____ Calcul Resultat pr l'annee ____
		for i = BaseYearB+1 to MaxYear
					VarText = "'Escalation '" & EscalationFmB
					response.write "<td title='Escalation = "&replace(EscalationFmB,";"," ")&"' >"
			
		' ____ Récupération des valeurs des indices dans la formule ____
			Tableau = split(EscalationFmB,";")
		
			Set RsIndicesValues = connection.execute("select * from Adv_Indices_Values as IV, " & _
			"Adv_Indices_Month as IM where IV.IdIndice = IM.IdIndice and iDate = IndiceDate AND IdBusCase = '" & _
			IdBusCase& "' AND Status = 'B' AND iYear = "&i)
	
			for z = 0 to UBound(Tableau, 1) -1 	'Pr chq elt de la FmEscal
			if Tableau(z)="*" OR Tableau(z)=":" OR Tableau(z)="+" OR Tableau(z)="-" OR Tableau(z) = "("  OR Tableau(z) = ")" then
				'Filtre les opérateurs
			elseif Tableau(z)="P0" OR (Tableau(z)="P(n-1)" AND i=BaseYearB) OR (Tableau(z)="P(n-2)" AND i<BaseYearB+2) then ' Si prix de base
				Tableau(z)=P0b
			elseif Tableau(z)="P(n-1)" then
				Tableau(z)=Pb(i-BaseYearB-1,TabInd)
			elseif Tableau(z)="P(n-2)" then
				Tableau(z)=Pb(i-BaseYearB-2,TabInd)
			elseif NOT(isnumeric(Tableau(z))) then ' Si indice trouvé dans formule
				if RsIndicesValues.eof then 'si pas d'indice
					response.write "Warn: No indice for " & Tableau(z) & "<br>"
					Tableau(z)=1
				else
					RsIndicesValues.movefirst
					do until RsIndicesValues.eof
						Indice=trim(RsIndicesValues("IdIndice")) & replace(trim(RsIndicesValues("IndicePos")),"9999","")
						if Tableau(z)=Indice then
							Tableau(z)=RsIndicesValues("iValue")
							exit do
						end if
						RsIndicesValues.movenext
					loop
				end if
			end if
			next
'				response.write EscalationFmB & "<br>"
'			for z1 = 0 to UBound(Tableau, 1)
'				response.write " " & Tableau(z1)
'			next
'			response.write "<br>"
		
			CalculTab Tableau ' Effectue les calculs dans Tableau(0)
	
			' ____ Récup prix normal escalé
			Pb(i-BaseYearB,TabInd)=Tableau(0)
			response.write Pb(i-BaseYearB,TabInd)
			response.write "</td>"
		next ' Boucle Annee

	case 3 'Theorical Price After action _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">B</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
		
			response.write "<td>"
			P(0,TabInd)=P0s
			response.write "Result "& P0s
			response.write "</td>"
			
		' ____ Calcul Resultat pr l'annee ____
		for i = BaseYearS+1 to MaxYear
		' ____ Verifie si si Freeze / JumpBack / reduction ____
			RsPricesRules.movefirst
			do until RsPricesRules.eof
				if RsPricesRules("iYear")=i then
					if RsPricesRules("JumpBack")= "x" then
						response.write "<td bgcolor=""#7AC7B1""> JumpB "
					elseif RsPricesRules("Freeze")= "x" then
						response.write "<td bgcolor=""#FF9966""> Friz "
					elseif RsPricesRules("PriceReduction")<>0 then
						response.write "<td bgcolor=""#FFFF99"">Price red (" & RsPricesRules("PriceReduction") &") <br>"
					else
					response.write "<td>"
					end if
					exit do
				end if
				RsPricesRules.movenext
			loop
			
		' ____ Récupération des valeurs des indices dans la formule ____
			Tableau = split(EscalationFmS,";")
		
			Set RsIndicesValues = connection.execute("select * from Adv_Indices_Values as IV, " & _
			"Adv_Indices_Month as IM where IV.IdIndice = IM.IdIndice and iDate = IndiceDate AND IdBusCase = '" & _
			IdBusCase& "' AND Status = 'S' AND iYear = "&i)
		
			for z = 0 to UBound(Tableau, 1) -1 	'Pr chq elt de la FmEscal
			if Tableau(z)="*" OR Tableau(z)=":" OR Tableau(z)="+" OR Tableau(z)="-" OR Tableau(z) = "("  OR Tableau(z) = ")" then
				'Filtre les opérateurs
				elseif Tableau(z)="P0" then ' Si prix de base
					Tableau(z)=P0s
				elseif Tableau(z)="P(n-1)" then
					if P(i-BaseYearS,TabInd)="" then P(i-BaseYearS,TabInd) = P0s
					Tableau(z)=P(i-BaseYearS,TabInd)
				elseif Tableau(z)="P(n-2)" then
					if i<BaseYearS-1 then
						Tableau(z)=P0s
					else
						Tableau(z)=P(i-BaseYearS,TabInd)
					end if
				elseif NOT(isnumeric(Tableau(z))) then ' Si indice trouvé dans formule
					if RsIndicesValues.eof then 'si pas d'indice
						Tableau(z)=1
						response.write "Warn: No indice for " & i & "<br>"
					else
						RsIndicesValues.movefirst
						do until RsIndicesValues.eof
							Indice=trim(RsIndicesValues("IdIndice")) & replace(trim(RsIndicesValues("IndicePos")),"9999","")
							if Tableau(z)=Indice then
								Tableau(z)=RsIndicesValues("iValue")
								exit do
							end if
							RsIndicesValues.movenext
						loop
					end if
			end if
			next
				
					for z1 = 0 to UBound(Tableau, 1)
							response.write " " &Tableau(z1)
					next		
				response.write "<br>"
			
			CalculTab Tableau ' Effectue les calculs dans Tableau(0)
	
			RsPricesRules.movefirst
			TabPricesRules=RsPricesRules.GetRows() ' récup Rs dans Tab
			
			' ____ Récup Pn si Freeze / JumpBack / reduction ____
			z1=0
			do while z1 < UBound(TabPricesRules, 2) and z1<500
				if TabPricesRules(1,z1)=i then 'année
					if TabPricesRules(3,z1)="x" then ' si Freeze
						if TabPricesRules(4,z1)="x" then 'si JumpBack stok px escalé ds N+1
							P(i-BaseYearS+1,TabInd)=Tableau(0)
							P(i-BaseYearS,TabInd)=P(i-BaseYearS-1,TabInd)
							response.write "JB "
						else
							P(i-BaseYearS+1,TabInd)=P(i-BaseYearS,TabInd)
							P(i-BaseYearS,TabInd)=P(i-BaseYearS-1,TabInd)
							response.write "Freezz "
						end if
					else 'ni Freeze ni JumpBack
						P(i-BaseYearS,TabInd)=Tableau(0)*(1+TabPricesRules(2,z1)*0.01)
						P(i-BaseYearS+1,TabInd)=P(i-BaseYearS,TabInd)
					end if
						response.write "Result "& P(i-BaseYearS,TabInd) & "<br>"
						response.write "Escal stockée "& P(i-BaseYearS+1,TabInd)

					exit do
				end if
				z1=z1+1
			loop
			response.write "</td>"
		next ' Boucle Annee

	case 4 'Qty of shipset from variant delivery schedule _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ____________________

		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">q</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"

			' ____ Calcul Resultat pr l'annee ____

		set RsBCVarQty =connection.Execute("SELECT SUM(Var_Qty) AS SumVarQty, IdVariant FROM dbo.Adv_Index_BC_Variant WHERE (IdBusCase = '"& IdBusCase &"') AND (TabInd = "& TabInd &") AND (Status = 'S') GROUP BY TabInd, IdVariant")
		set RsVariant =connection.Execute("SELECT YEAR(dbo.Adv_DelivSchedule.iDate) AS CurrentYear, dbo.Adv_Index_BC_Variant.IdVariant, SUM(dbo.Adv_DelivSchedule.Qty) AS SommeDeQty FROM Adv_DelivSchedule, Adv_Index_BC_Variant WHERE Adv_DelivSchedule.IdVariant = Adv_Index_BC_Variant.IdVariant AND IdBusCase = '"& IdBusCase &"' AND TabInd = "& TabInd &" AND Status = 'S' GROUP BY YEAR(dbo.Adv_DelivSchedule.iDate), dbo.Adv_Index_BC_Variant.IdVariant ORDER BY YEAR(dbo.Adv_DelivSchedule.iDate)")
		TabBCVarQty=RsBCVarQty.GetRows()
		TabVariant=RsVariant.GetRows()

		for i = BaseYearS to MaxYear
			response.write "<td> &nbsp;"
			z1=0
			response.write "DelivSced Qty ("&i& ") * BC shipset Qty<br>"
			do while z1 <= UBound(TabVariant, 2) and z1<500
				if TabVariant(0,z1)=i then 'année
					z2=0
					do while z2 <= UBound(TabBCVarQty, 2) and z2<500
						if TabVariant(1,z1)=TabBCVarQty(1,z2) then
							if z2>0 then response.write " + "
							response.write TabVariant(2,z1) & " * "&TabBCVarQty(0,z2) & "<br>"
							DvScQty(i-BaseYearS,TabInd)=DvScQty(i-BaseYearS,TabInd)+ (TabVariant(2,z1) * TabBCVarQty(0,z2))
						end if
						z2=z2+1
					loop
				end if
				z1=z1+1
			loop
			response.write " = " & DvScQty(i-BaseYearS,TabInd)
			response.write "</td>"
		next 
			
	case 5 'Theorical saving based on variant BC calculation _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">(B-A)*q</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"

			' ____ Calcul Resultat pr l'annee ____
		for i = BaseYearS to MaxYear
			response.write "<td bgcolor=""#FFCC99"">" & _
			(P(i-BaseYearS,TabInd)-Pb(i-BaseYearS,TabInd))*DvScQty(i-BaseYearS,TabInd)
			response.write "</td>"
		next
		response.write "</td>"
		
	case 6 'SAP Negociated Price _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">B'</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"

			set RsSAP =connection.Execute("SELECT SUM(DISTINCT dbo.Adv_SapArticle.NegoPrice) AS PxSAP, dbo.Adv_SapArticle.iYear FROM dbo.Adv_SapArticle INNER JOIN dbo.Adv_Index_BC_Article ON dbo.Adv_SapArticle.IdArticle = dbo.Adv_Index_BC_Article.IdArticle WHERE (dbo.Adv_Index_BC_Article.TabInd = "& TabInd &") AND (dbo.Adv_SapArticle.NegoPrice > 0) AND (dbo.Adv_Index_BC_Article.IdBusCase = '"& IdBusCase &"') AND (dbo.Adv_Index_BC_Article.Status = 'S') GROUP BY dbo.Adv_SapArticle.iYear")

			' ____ Calcul Resultat pr l'annee ____
		for i = BaseYearS to MaxYear
			response.write "<td bgcolor=""#FFCC99"">"
			if not(RsSAP.eof) then
				RsSAP.movefirst
				do until RsSAP.eof
					if RsSAP("iYear")=i then ' ____ Affiche PxSAP ____
						PxSAP(i-BaseYearS,TabInd)=Cdbl(RsSAP("PxSAP"))*DvScQty(i-BaseYearS,TabInd)
						response.write PxSAP(i-BaseYearS,TabInd)
						exit do
					end if
					RsSAP.movenext
				loop
			else
				response.write "No record"
			end if
			response.write "</td>"
		next
		
	case 7 'Negociated price monitoring against theorical price _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">B'-B</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"

			' ____ Calcul Resultat pr l'annee ____
		for i = BaseYearS to MaxYear
			response.write "<td bgcolor=""#FFCC99"">" & _
			PxSAP(i-BaseYearS,TabInd)-P(i-BaseYearS,TabInd)
			response.write "</td>"
		next
		response.write "</td>"

	case 8 'Theorical saving based on SAP Negociated Price _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">(B'-A)*q</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
			
			' ____ Calcul Resultat pr l'annee ____
		for i = BaseYearS to MaxYear
			response.write "<td bgcolor=""#FFCC99"">" & _
			(PxSAP(i-BaseYearS,TabInd)-Pb(i-BaseYearS,TabInd))*DvScQty(i-BaseYearS,TabInd)
			response.write "</td>"
		next
		response.write "</td>"

	case 9 'Good receipt price per shipset in SAP _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">C</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"

		set RsAmountSAP =connection.Execute("SELECT SUM(DISTINCT dbo.Adv_SapArticle.SCLAmount) AS AmountSAP, dbo.Adv_SapArticle.iYear FROM dbo.Adv_SapArticle INNER JOIN dbo.Adv_Index_BC_Article ON dbo.Adv_SapArticle.IdArticle = dbo.Adv_Index_BC_Article.IdArticle WHERE (dbo.Adv_Index_BC_Article.TabInd = "& TabInd &") AND (dbo.Adv_SapArticle.SCLAmount > 0) AND (dbo.Adv_Index_BC_Article.IdBusCase = '"& IdBusCase &"') AND (dbo.Adv_Index_BC_Article.Status = 'S') GROUP BY dbo.Adv_SapArticle.iYear")

		do until RsAmountSAP.eof
			for i = BaseYearS to MaxYear
				if trim(RsAmountSAP("iYear"))=cstr(i) then
					AmountSAP(i-BaseYearS,TabInd)=Cdbl(RsAmountSAP("AmountSAP"))
					exit for
				end if
			Next
			RsAmountSAP.movenext
		loop
		for i = BaseYearS to MaxYear
			if AmountSAP(i-BaseYearS,TabInd)="" then
				response.write "<td align=""center"">No Record</td>"
			else
				response.write "<td align=""center"">" & AmountSAP(i-BaseYearS,TabInd) & "</td>"
			end if
		Next

	case 10 'Delta Good receipt price per shipset and theorical price _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">C-B</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
			
		for i = BaseYearS to MaxYear
			response.write "<td align=""center"">" & AmountSAP(i-BaseYearS,TabInd)-P(i-BaseYearS,TabInd) & "</td>"
		Next


		case else
		end select
		
		response.write "</tr>"
	next
	
	if not(RsArticles.eof) then
		TabInd=RsArticles("TabInd")
	else
		TabInd=0
	end if
	response.write "</table>"
	
loop '_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_____
'__________________________________________________________________________________________


	response.write	"<table border=""2"" width=""100%"" style=""border: 3px ridge #CCCCFF"">" & _
	"<tr><td height=""5"" align=""center"" colspan=""5"">CONSOLIDATED REPORT</td></tr>" & _
	"<tr>" & _
		"<td nowrap width=""180"" height=""5"" colspan=""2"" style=""background-color: #B0CFD0;"" align=""center"">Year</td>"
	for i = BaseYearS to MaxYear
		response.write	"<td nowrap width =""100"" style=""background-color: #B0CFD0"" align=""center"">"& _
		i & "</td>"
	next

	for n=1 to 6 'nombre de Lignes
		
		Select case n
		case 1
		Name="Qty of shipset from variant delivery schedule"
		case 2
		Name="Theorical saving based on variant BC calculation"
		case 3
		Name="Theorical saving based on SAP Negociated Price"
		case 4
		Name="Initial ADVANCE secured savings"
		case 5
		Name="Delta secured savings ON theorical realised"
		case 6
		Name="savings realised based on GR price in SAP"
		end select
		
		Select case n

	case 1 'Qty of shipset from variant delivery schedule _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">q</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
			
		for i = BaseYearS to MaxYear
			VarSomme=0
			For TabInd = 0 to UBound(Pb,2)
				if Pb(0,TabInd)<>"" then
					VarSomme=VarSomme+DvScQty(i-BaseYearS,TabInd)
				end if
			next
			response.write "<td align=""center"">" & VarSomme & "</td>"
		Next
		response.write	"</td>"			

	case 2 'Theorical saving based on variant BC calculation_ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">(B-A)*q</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
			
		for i = BaseYearS to MaxYear
			VarSomme=0
			For TabInd = 0 to UBound(Pb,2)
				if Pb(0,TabInd)<>"" then
					VarSomme=VarSomme+(P(i-BaseYearS,TabInd)-Pb(i-BaseYearS,TabInd))*DvScQty(i-BaseYearS,TabInd)
				end if
			next
			response.write "<td align=""center"">" & VarSomme & "</td>"
		Next
		response.write	"</td>"			

	case 3 'Theorical saving based on SAP Negociated Price _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">(B'-A)*q</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
			
		for i = BaseYearS to MaxYear
			VarSomme=0
			For TabInd = 0 to UBound(Pb,2)
				if Pb(0,TabInd)<>"" then
					VarSomme=VarSomme+(PxSAP(i-BaseYearS,TabInd)-Pb(i-BaseYearS,TabInd))*DvScQty(i-BaseYearS,TabInd)
				end if
			next
			response.write "<td align=""center"">" & VarSomme & "</td>"
		Next
		response.write	"</td>"

	case 4 'Initial ADVANCE secured savings _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">D</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"

		for i = BaseYearS to MaxYear
			response.write "<td align=""center"">"
			response.write TabPricesRules(5,i-BaseYearS)
			response.write	"</td>"
		next


	case 5 'Delta secured savings / theorical realised _ _ _ _ _ _ _ _ _ _ ____________________
		
		response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">((B'-A)*q)-((B-A)*q)</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
			
		for i = BaseYearS to MaxYear
			VarSomme=0
			For TabInd = 0 to UBound(Pb,2)
				if Pb(0,TabInd)<>"" then
					VarSomme=VarSomme+(PxSAP(i-BaseYearS,TabInd)-Pb(i-BaseYearS,TabInd))*DvScQty(i-BaseYearS,TabInd) - (P(i-BaseYearS,TabInd)-Pb(i-BaseYearS,TabInd))*DvScQty(i-BaseYearS,TabInd)
				end if
			next
			if VarSomme>0 then
				response.write "<td align=""center"">" & VarSomme & "</td>"
			else
				response.write "<td align=""center"" style=""color: #CC0000"">" & VarSomme & "</td>"
			end if
		Next
		response.write	"</td>"

	case 6 'savings realised based on GR price in SAP _ _ _ _ _ _ _ _ _ _ ____________________
		
	response.write	"<tr>" & _
			"<td nowrap width=""40"" height=""5"" align=""center"">(C-A)*q</td>"& _
			"<td nowrap width=""140"">"& Name &"</td>"
			
		for i = BaseYearS to MaxYear
			VarSomme=0
			For TabInd = 0 to UBound(Pb,2)
				if Pb(0,TabInd)<>"" then
					VarSomme=VarSomme+(AmountSAP(i-BaseYearS,TabInd)-Pb(i-BaseYearS,TabInd))*DvScQty(i-BaseYearS,TabInd)
				end if
			next
			response.write "<td align=""center"">" & VarSomme & "</td>"
		Next
		response.write	"</td>"

		end select
	next

		response.write "</tr>"
	response.write "</table>"

		'----------teste si le fichier des commentaires existe------------
		Set FSO = Server.CreateObject("Scripting.FileSystemObject")
		NomFichier=Server.MapPath("Textes/"&IdBusCase&".txt")
		If FSO.FileExists(NomFichier) Then
			Set Fichier = FSO.OpenTextFile(NomFichier, 1, FALSE, FALSE)
			if not (Fichier.atEndOfStream) then Comments = Fichier.ReadAll
			set Fichier=nothing
		end if
		Set FSO = Nothing
		
		
		
		
		
'Legende
response.write "<table border=""1"" width=""30%"">" & _
	"<tr>" & _
		"<td nowrap width =""30%"" bgcolor=""#FF9966""></td>" & _
		"<td>Freeze</td>" & _
	"</tr>" & _
	"<tr>" & _
		"<td bgcolor=""#7AC7B1""></td>" & _
		"<td nowarp >Freeze with jumpback</td>" & _
	"</tr>" & _
	"<tr>" & _
		"<td bgcolor=""#FFFFCC""></td>" & _
		"<td>Price reduction</td>" & _
	"</tr>" & _
"</table>"

set RsSAPArticles =connection.Execute("SELECT dbo.Adv_SapArticle.IdArticle, dbo.Adv_SapArticle.iYear, dbo.Adv_SapArticle.PurchasingOrg, dbo.Adv_SapArticle.Plant, dbo.Adv_SapArticle.IdSupplier, dbo.Adv_SapArticle.SCLAmount, dbo.Adv_SapArticle.IdCurrency, dbo.Adv_SapArticle.SCLQty, dbo.Adv_Article.Designation, dbo.Adv_SapArticle.NegoPrice FROM dbo.Adv_SapArticle INNER JOIN dbo.Adv_Article ON dbo.Adv_SapArticle.IdArticle = dbo.Adv_Article.IdArticle INNER JOIN dbo.Adv_Index_BC_Article ON dbo.Adv_Article.IdArticle = dbo.Adv_Index_BC_Article.IdArticle WHERE (dbo.Adv_Index_BC_Article.IdBusCase = '"& IdBusCase &"') GROUP BY dbo.Adv_SapArticle.IdArticle, dbo.Adv_SapArticle.iYear, dbo.Adv_SapArticle.PurchasingOrg, dbo.Adv_SapArticle.Plant, dbo.Adv_SapArticle.IdSupplier, dbo.Adv_SapArticle.SCLAmount, dbo.Adv_SapArticle.IdCurrency, dbo.Adv_SapArticle.SCLQty, dbo.Adv_Article.Designation, dbo.Adv_SapArticle.NegoPrice ORDER BY dbo.Adv_SapArticle.IdArticle, dbo.Adv_SapArticle.Plant, dbo.Adv_SapArticle.IdSupplier")

TabSAPArticles=RsSAPArticles.GetRows() ' récup Rs dans Tab
ReDim TSAPAres(MaxYear-BaseYearS+3,0)

if UBound(TabSAPArticles, 2)>0 then
	response.write "<table border=""1"" width=""30%"">" & _
		"<tr><td colspan="& MaxYear-BaseYearS+3 &">For Information only</td></tr>"& _
		"<tr><td nowarp colspan="& MaxYear-BaseYearS+3 &">Quantities delivered in SAP</td></tr>"& _
		"<tr><td>Part Numbers</td>"& _
		"<td nowarp width=""90"">Designation</td>"
	for i = BaseYearS to MaxYear
		response.write "<td nowarp width=""30"">"&i&"</td>"
	next
		response.write "<td nowarp width=""40"">Total</td>"
	response.write "</tr>"
		
	z1=0
	j=0
	For z1 = 0 to UBound(TabSAPArticles, 2)-1
		if TabSAPArticles(3,z1)<>TabSAPArticles(3,z1+1) OR TabSAPArticles(4,z1)<>TabSAPArticles(4,z1+1) OR TabSAPArticles(0,z1)<>TabSAPArticles(0,z1+1) then
			TSAPAres(TabSAPArticles(1,z1)-BaseYearS+2,j)=TabSAPArticles(7,z1)
			TSAPAres(0,j)=TabSAPArticles(0,z1)
			TSAPAres(1,j)=TabSAPArticles(8,z1)
			j=j+1
			ReDim Preserve TSAPAres(MaxYear-BaseYearS+3,j)
		else
			TSAPAres(TabSAPArticles(1,z1)-BaseYearS+2,j)=TabSAPArticles(7,z1)		
		end if
	next
	
	if TabSAPArticles(3,z1)<>TabSAPArticles(3,z1-1) OR TabSAPArticles(4,z1)<>TabSAPArticles(4,z1-1) OR TabSAPArticles(0,z1)<>TabSAPArticles(0,z1-1) then
		j=j+1
		ReDim Preserve TSAPAres(MaxYear-BaseYearS+3,j)
		TSAPAres(0,j)=TabSAPArticles(0,z1)
		TSAPAres(1,j)=TabSAPArticles(8,z1)
	else
		TSAPAres(0,j)=TabSAPArticles(0,z1)
		TSAPAres(1,j)=TabSAPArticles(8,z1)
	end if
		TSAPAres(TabSAPArticles(1,z1)-BaseYearS+2,j)=TabSAPArticles(7,z1)		
	
	For z1 = 0 to UBound(TSAPAres, 2)
		Total=0
		response.write "<tr>"
		for i = 0 to UBound(TSAPAres, 1)-1
			response.write "<td nowarp width=""30"">"
			if TSAPAres(i,z1)<>"" then
				response.write TSAPAres(i,z1)
				if isnumeric(TSAPAres(i,z1)) then Total=Total+TSAPAres(i,z1)
			else
				response.write "-"
			end if
			response.write "</td>"
		next
		response.write "<td nowarp width=""30"">" & Total & "</td>"
		response.write "</tr>"
	next
	response.write "</table>"
end if
		
		
		
%>
<table border="2" width="100%" style="border-style:ridge; border-width:3px; " bordercolorlight="#00D2A8" bordercolordark="#00A882">
<tr><td height="5" align="center" width="54%">GAP Analysis <SMALL>(Dble Click to edit)</SMALL></td></tr>
<tr><td colspan="2" align="center"><textarea ondblclick="OuvrirFenetre ('Texte_ASRaM.asp?','GAP Analysis','target=\'mainFrame\'')" cols="118" name="Comments" readonly rows="7" style="border-style: inset; border-width: 2px; color: #008080">
<%=Comments%></textarea>
</td></tr>
</table>
</form>
</body>
<script language="JavaScript">
<!--#INCLUDE FILE="connection/deconnection.asp"-->
</script>
</html>
