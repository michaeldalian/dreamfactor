<html>
<script language="javascript">
<!--#INCLUDE FILE="connection/connection.asp"-->

var onglet = ""

//Changement de style onMouseover
function change(divers){
divers.style.color="#0C0C5E"; //couleur de surbrillance de police
divers.style.backgroundColor="#DCE9F6"; //couleur de surbrillance de l'arrière plan
}
//Changement de style onMouseout (retour à la couleur d'origine)
function change_back(divers){
divers.style.color="#CCCCFF"; //couleur d'origine de police
divers.style.backgroundColor=""; //couleur d'origine de l'arrière plan
}

function ValideFm()
{
if(window.event.keyCode==13) 
{
MainFile.submit()
}}

function estDate(dateStr) 
{
      var datePat = /^(\d{1,2})(\/)(\d{1,2})(\/)(\d{4})$/;
      var matchArray = dateStr.match(datePat); // is the format ok?

        if (matchArray == null) 
        {
               alert("Please enter date as dd/mm/yyyy .");
               return false;
        }
        month = matchArray[3]; // p@rse date into variables
        day = matchArray[1];
        year = matchArray[5];
 
        if (month < 1 || month > 12) 
        { // check month range
                alert("Month must be between 1 and 12.");
               return false;
        }
 
        if (day < 1 || day > 31) 
        {
                alert("Day must be between 1 and 31.");
               return false;
        }
 
        if ((month==4 || month==6 || month==9 || month==11) && day==31) 
        {
               alert("Month "+month+" doesn`t have 31 days!")
               return false;
        }
 
        if (month == 2) 
        { // check for february 29th
               var isleap = (year % 4 == 0 && (year % 100 != 0 || year % 400 == 0));
               if (day > 29 || (day==29 && !isleap)) 
               {
                       alert("February " + year + " doesn`t have " + day + " days!");
                       return false;
               }
        }
        return true; // date is valid
}


function TestChp(Nom){// Procedure de validation du formulaire
	var str2 = document.getElementById('t_' + Nom.substring(2)).value
  // On teste si les éléments du formulaire sont ok (valeurs...)
  re = /^(\d\d?)(\/|-|\.)(\d\d?)(\/|-|\.)(\d\d)(\d\d)?$/;
 return (chaine.search(re) != -1);
  
  if(!chaine)
  {
    alert("It's not a Date format")
  document.getElementById(Nom).value = ''
    // et on indique de ne pas envoyer le formulaire
    return false;}
}
function Valiste(Liste, Champ){// Procedure de validation alert(Nom)
	document.getElementById(Champ).value = document.getElementById(Liste).options[document.getElementById(Liste).selectedIndex].text
}

function CtrlNom(Nom){// Procedure de validation
	if (document.getElementById(Nom).value == "")
{
alert("Renseigner le nom!!!")
window.event.keyCode = ""
return false;
}
}

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

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>ASRaM</title>
<script language="JavaScript">
<!--
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
// -->
</script>
</head>
<%
'affiche toute les valeurs envoyée avec la methode get 
'  For Each chaine_requete In Request.QueryString
' 		Response.Write "<tr>;<td>" & chaine_requete & "</td> = <td>" & Request.QueryString(chaine_requete) & "</td></tr>" 
'  Next

IdBusCaseNew = request.querystring("IdBusCaseNew")
IdBusCase = request.querystring("IdBusCase")
CreateFirst= request.querystring("CreateFirst")
onglet = request.querystring("onglet")

if IdBusCaseNew <> "" then 'si nouveau
  'controle s'il exist déjà alors que l'utilisateur clic sur create
  	set RsBusCaseTestIfExist =connection.execute("select * from Adv_BusCase where IdBusCase = '"&IdBusCaseNew&"' order by IdBusCase")
 	if not (RsBusCaseTestIfExist.eof and RsBusCaseTestIfExist.bof) then 'Si Existant
		IdBusCaseNew = ""
		%>
		 <font color="#CC0000">Attention: Business Case existant!</font>
	<%else
		CreationDate = date()
		SecureDate = date()
		Status="B"
		ps_adv_AddBusCase = "ps_adv_AddBusCase '" &IdBusCaseNew& "','" &IdBuyer& "','" & _
		IdALStation& "','" &Title& "','" &Designation& "','" &formatDate(CreationDate)& "','" & _
		formatDate(SecureDate)& "','" &IdCreator& "','" &Status& "','" &Plant& "','" &PurchasingOrg& "','" &Comments& "'"
		response.write ps_adv_AddBusCase
		set AddProc =connection.Execute(ps_adv_AddBusCase)
	end if
	IdBusCase = request.querystring("IdBusCaseNew")
end if

if IdBusCase = "" then%>
<form name="first_buscase">
	Enter a new Business Case number : <input type="text" name="IdBusCaseNew" size="14" value ="">
	<input type="submit" value="Envoyer" name="CreateFirst">
</form>
<%
response.end
end if

' Supprimer BusCase ______________________________________________________________
if Request.querystring("BtDelete") = "Delete" then 
	IdBusCase = trim(Request.querystring("IdBusCase"))
	if IdBusCase <> "" then
	
	'Index_BC_Article
	set RsArticle =connection.Execute("select * from Adv_Index_BC_Article where IdBusCase = '" & IdBusCase & "'")
	do until RsArticle.eof
			IdLine=RsArticle("IdLine")
			IdArticle = trim(RsArticle("IdArticle"))
			TabInd = trim(RsArticle("TabInd"))
			Status = trim(RsArticle("Status"))

			ps_adv_MajIndex_BC_Article = "ps_adv_MajIndex_BC_Article '"&IdBusCase&"','"& _
			IdLine&"','"&IdArticle&"','"&TabInd&"','"&Status&"','','','','','','oui'"
			response.write "MAJ" & ps_adv_MajIndex_BC_Article
			set MajArticle=connection.Execute(ps_adv_MajIndex_BC_Article)	
		RsArticle.movenext
	loop

'PricesRule
	Set RsPricesRules =connection.execute("select * from Adv_PricesRules where IdBusCase = '"&IdBusCase&"'")
	
	do until RsPricesRules.eof
		iYear=RsPricesRules("iYear")
	
		ps_adv_MajPricesRules = "ps_adv_MajPricesRules '"&IdBusCase&"','"&iYear&"','','','','','oui'"
		Response.write ps_adv_MajPricesRules
		set AddProc =connection.Execute(ps_adv_MajPricesRules)
		RsPricesRules.movenext
	loop

'Indices_Month
	set RsIndices =connection.execute("select * from Adv_Indices_month where IdBusCase = '"&IdBusCase&"'")

	do until RsIndices.eof

		Status=RsIndices("Status")
		IdIndice=RsIndices("IdIndice")
		IndicePos=RsIndices("IndicePos")
		iYear=RsIndices("iYear")
			
		ps_adv_MajIndices_Month = "ps_adv_MajIndices_Month '" &IdBusCase& "','" &Status& "','" & _
	 	IdIndice& "','" &IndicePos& "','" &iYear& "','','oui'"
		Response.write ps_adv_MajIndices_Month
		set AddProc =connection.Execute(ps_adv_MajIndices_Month)
		RsIndices.movenext
	loop

'BusCase
	ps_adv_DelBusCase = "ps_adv_DelBusCase '" &IdBusCase& "'"
	response.write ps_adv_DelBusCase
	set AddProc =connection.Execute(ps_adv_DelBusCase)

	end if
	response.end
end if


set RsAssemblyLine =connection.execute("select * from Adv_AssemblyLine order by IdALStation")
set RsCurrency =connection.execute("select * from Adv_Currency")
set RsUnit =connection.execute("select * from Adv_Unit")
set RsBuyer =connection.execute("select * from Adv_Buyer order by IdBuyer")
set RsPlantSAP =connection.execute("SELECT dbo.Adv_SapArticle.Plant FROM dbo.Adv_SapArticle INNER JOIN dbo.Adv_Index_BC_Article ON dbo.Adv_SapArticle.IdArticle = dbo.Adv_Index_BC_Article.IdArticle GROUP BY dbo.Adv_SapArticle.Plant, dbo.Adv_Index_BC_Article.IdBusCase")
set RsUnitOrgSAP =connection.execute("SELECT dbo.Adv_SapArticle.PurchasingOrg FROM dbo.Adv_SapArticle INNER JOIN dbo.Adv_Index_BC_Article ON dbo.Adv_SapArticle.IdArticle = dbo.Adv_Index_BC_Article.IdArticle GROUP BY dbo.Adv_Index_BC_Article.IdBusCase, dbo.Adv_SapArticle.PurchasingOrg")

%>
<body onload="FP_preloadImgs(/*url*/'images/buttonEdit.jpg', /*url*/'images/buttonB7.jpg', /*url*/'images/button1.jpg', /*url*/'images/button2.jpg', /*url*/'images/button177.jpg', /*url*/'images/button178.jpg')">
<!-- Cree un form: focus sur controle et valide si on presse enter -->
<form name="MainFile" action="BusEdit_ASRaM.asp" method="get" onkeydown ="ValideFm();">

<%' Sauver vers BDD ______________________________________________________________
if Request.querystring("BSave") = "Save" then 

	IdBusCase = trim(Request.querystring("IdBusCase"))
	IdBuyer = trim(Request.querystring("IdBuyer"))
	if IsNull(IdBuyer) then IdBuyer = ""
	IdALStation = trim(Request.querystring("IdALStation"))
	Plant = trim(Request.querystring("Plant"))
	PurchasingOrg = trim(Request.querystring("PurchasingOrg"))
	Title = trim(Request.querystring("Title"))
	Designation = trim(Request.querystring("Designation"))
	CreationDate = trim(Request.querystring("CreationDate"))
	SecureDate = trim(Request.querystring("SecureDate"))
	IdCreator = trim(Request.querystring("IdCreator"))
	Status = trim(Request.querystring("Status"))
	Comments= trim(Request.querystring("Comments"))
	
	ps_adv_AddBusCase = "ps_adv_AddBusCase '" &IdBusCase& "','" &IdBuyer& "','" & _
	 IdALStation& "','" &Title& "','" &Designation& "','" &formatDate(CreationDate)& "','" & _
	formatDate(SecureDate)& "','" &IdCreator& "','" &Status& "','" &Plant& "','" &PurchasingOrg& "','" &Comments& "'"
	response.write ps_adv_AddBusCase
	set AddProc =connection.Execute(ps_adv_AddBusCase)

elseif Request.querystring("onglet") = "Lock and Save" then

	IdBusCase = Request.querystring("IdBusCase")
	IdBuyer = Request.querystring("IdBuyer")
	if IsNull(IdBuyer) then IdBuyer = ""
	IdALStation = Request.querystring("IdALStation")
	Plant = trim(Request.querystring("Plant"))
	PurchasingOrg = trim(Request.querystring("PurchasingOrg"))
	Title = Request.querystring("Title")
	Designation = Request.querystring("Designation")
	CreationDate = Request.querystring("CreationDate")
	SecureDate = Request.querystring("SecureDate")
	IdCreator = Request.querystring("IdCreator")
	Status = "S"
	Comments= trim(Request.querystring("Comments"))
	
	ps_adv_AddBusCase = "ps_adv_AddBusCase '" &IdBusCase& "','" &IdBuyer& "','" & _
	 IdALStation& "','" &Title& "','" &Designation& "','" &formatDate(CreationDate)& "','" & _
	formatDate(SecureDate)& "','" &IdCreator& "','S','" &Plant& "','" &PurchasingOrg& "','" &Comments& "'"
	set AddProc =connection.Execute(ps_adv_AddBusCase)
	response.write ps_adv_AddBusCase
	set RsArticle =connection.Execute("select * from Adv_Article, Adv_Index_BC_Article where Adv_Article.IdArticle=Adv_Index_BC_Article.IdArticle AND IdBusCase = '" & IdBusCase & "' AND Status = 'B' order by TabInd, IdLine, ValidTo desc")
	
		do until RsArticle.eof
			IdLine=RsArticle("IdLine")
			IdArticle = trim(RsArticle("IdArticle"))
			TabInd = trim(RsArticle("TabInd"))
			ValidTo = RsArticle("ValidTo")
			Art_Qty = replace (RsArticle("Art_Qty"),",",".")
			Art_Price = replace (RsArticle("Art_Price"),",",".")
			IdUnit = RsArticle("IdUnit")
			IdCurrency=RsArticle("IdCurrency")
		
			if IdArticle<>"" then
				ps_adv_MajIndex_BC_Article = "ps_adv_MajIndex_BC_Article '"&IdBusCase&"','"& _
				IdLine&"','"&IdArticle&"','"&TabInd&"','S','"&ValidTo&"','"& _
				Art_Qty&"','"&Art_Price&"','"&IdUnit&"','"&IdCurrency&"',''"
				set MajArticle=connection.Execute(ps_adv_MajIndex_BC_Article)
			end if
			RsArticle.movenext
		loop
		
	set RsVariant = CONNECTION.execute("select * from Adv_Variant, Adv_Index_BC_Variant where Adv_Variant.IdVariant=Adv_Index_BC_Variant.IdVariant AND IdBusCase = '" & IdBusCase & "' AND Status = 'B' order by TabInd")
	
		do until RsVariant.eof
			IdVariant = trim(RsVariant("IdVariant"))
			TabInd = trim(RsVariant("TabInd"))
			Var_Qty = trim(RsVariant("Var_Qty"))
		
			if IdVariant<>"" then 
				ps_adv_MajIndex_BC_Variant = "ps_adv_MajIndex_BC_Variant '"&IdBusCase&"','"& _
				IdVariant&"','S','"&TabInd&"','"&Var_Qty&"','"&IdMacro&"',''"
				set MajVariant=connection.Execute(ps_adv_MajIndex_BC_Variant)
			end if
			RsVariant.movenext
		loop
		
	set RsEscal =connection.execute("select * from Adv_Escalation where IdBusCase = '" & _
	IdBusCase & "' and Status = 'B'")
 	if not (RsEscal.eof and RsEscal.bof) then ' Fmul existe
 		RsEscal.movefirst
		EscalationFm = trim(RsEscal("EscalationFm"))
		iYear = RsEscal("iYear")
		else
		iYear = 2003
	end if
		
		ps_Adv_AddEscalation = "ps_Adv_AddEscalation '" & IdBusCase & "','S','" & EscalationFm & "','" & iYear & "'"
		set Maj=connection.Execute(ps_Adv_AddEscalation)
		
end if

  if IdBusCase <> "" and IdBusCaseNew = "" then 'edit existant

	set RsBusCase =connection.execute("select * from Adv_BusCase where IdBusCase = '" & IdBusCase & "' order by IdBusCase")
	
	RsBusCase.movefirst
		IdBusCase = trim(RsBusCase("IdBusCase"))
		IdBuyer = trim(RsBusCase("IdBuyer"))
		IdALStation = trim(RsBusCase("IdALStation"))
		Plant = trim(Request.querystring("Plant"))
		PurchasingOrg = trim(Request.querystring("PurchasingOrg"))
		Title = trim(RsBusCase("Title"))
		Designation = trim(RsBusCase("Designation"))
		CreationDate = trim(RsBusCase("CreationDate"))
		SecureDate = trim(RsBusCase("SecureDate"))
		IdCreator = trim(RsBusCase("IdCreator"))
		Status = trim(RsBusCase("Status"))
		Comments= trim(RsBusCase("Comments"))
		
end if%>

	<!-- Cree un tableau -->
	<table border="0" width="100%" id="table1" style="border: 2px ridge #0000FF; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bordercolorlight="#00FFFF" bordercolordark="#0000FF">
	<td width="15%" height="41"> Creator:  
	<td width="14%" height="41">
	<input type="text" style="border-style: solid; border-width: 1px; " name="IdCreator" size="14" value="<%=IdCreator%>"></td>
	<td height="41" width="15%" >Business Case Title :</td>
	<td width="27%" height="41">
	<textarea name="Title" style="border-style: inset; border-width: 2px" cols="22" rows="2" ><%=Title%></textarea>
	</td>
	<td width="11%" height="41" align="center"> Id Number :</td>
	<td height="41">
	<input type="text" name="IdBusCase" size="14" value ="<%=IdBusCase%>" readonly style="border: 2px ridge #0000FF">
	</td>
	<tr><!-- Ligne suivante ---------------------------------->
	<td width="15%"> Creation Date:<td width="14%">
	<input type="text" readonly name="CreationDate"  size="14"
	value ="<%=CreationDate%>" style="border-style: solid; border-width: 1px"></td>
	<td width="15%" >Status</td>
	<td style="border:1px solid #000000; " bordercolor="#FFFFFF">
	<%if Status = "S" then
		response.write "Secured"
	else
		response.write "BaseLine"
	end if %>
	</td>
	<td>
	<a href="SupplierSelect_ASRaM.asp?IdBusCase=<%=IdBusCase%>" target=_blank width="300" height="300">
	suppliers<IMG border=0  onclick= "document.getElementById('BSave').click()" height=16 hspace=5 src="images/Create.JPG" width=16></a> 
	<td>
	<%=IdSupplier%>
	<tr><!-- Ligne suivante ---------------------------------->
	<td width="15%" height="35"> Assembly line Station :<td width="14%" height="35">
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
	<td height="70" width="15%" rowspan="2">Business Case Designation </td>
	<td width="27%" height="70" rowspan="2">
	<textarea name="Designation" style="border-style: inset; border-width: 2px" cols="23" rows="4"><%=trim(Designation)%></textarea></td>
	<td  width="26%" colspan="2" rowspan="5">

	<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF" id="table4">
		<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">IdSupplier</td>
		<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%">Designation</td>
		<%
		set RsSupplier =connection.execute("select * from Adv_Index_BC_Supplier as BCS, Adv_Supplier as Sup" & _ 
		" where BCS.IdSupplier = Sup.IdSupplier and BCS.IdBusCase = '" & _
		IdBusCase & "' order by BCS.IdSupplier")
		if not (RsSupplier.eof and RsSupplier.bof) then
			do until RsSupplier.eof %>
			<tr>
			<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px; float:left" bgcolor="#FFFFFF"><%=trim(RsSupplier("IdSupplier"))%></td>
			<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px; float:left" bgcolor="#FFFFFF" width="6%"><%=trim(RsSupplier("Designation"))%></td>
			</tr>
			<% RsSupplier.movenext
			Loop
			end if%>
	</table>				
	
	</td>
	<tr>
	<td nowrap width="15%" height="35"> Responsible buyer :<td width="14%" height="35">
	<select size="1" name="IdBuyer" >
	<OPTION></OPTION>
	<%RsBuyer.movefirst
	do until RsBuyer.eof
		if IdBuyer = trim(RsBuyer("IdBuyer")) then %>
			<OPTION Value ="<%=trim(RsBuyer("IdBuyer"))%>" selected>
			<%=RsBuyer("Designation")%></OPTION>
		<%else%>
			<OPTION Value ="<%=trim(RsBuyer("IdBuyer"))%>">
			<%=RsBuyer("Designation")%></OPTION>
		<%end if
		RsBuyer.movenext
	loop%>
	</SELECT></td>
	<tr><!-- Ligne suivante ---------------------------------->
	<td nowrap width="15%" height="20"> Plant :<td width="14%" height="20">
	<select size="1" name="Plant" >
	<OPTION></OPTION>
	<%RsPlantSAP.movefirst
	do until RsPlantSAP.eof
		if Plant = trim(RsPlantSAP("Plant")) then %>
			<OPTION Value ="<%=trim(RsPlantSAP("Plant"))%>" selected >
			<%=RsPlantSAP("Plant")%></OPTION>
		<%else%>
			<OPTION Value ="<%=trim(RsPlantSAP("Plant"))%>">
			<%=RsPlantSAP("Plant")%></OPTION>
		<%end if
		RsPlantSAP.movenext
	loop%>
	</SELECT></td>
	<td height="70" width="15%" rowspan="2" style="color: #008080">Comments / 
	Notes</td>
	<td width="27%" height="70" rowspan="2">
	<textarea name="Comments" cols="23" rows="4" style="border-style: inset; border-width: 2px; color: #008080"><%=trim(Comments)%></textarea></td>

	</td>
	<tr>
	<td width="15%" height="20"> Unit Organisation :<td width="14%" height="20">
	<select size="1" name="PurchasingOrg" >
	<OPTION></OPTION>
	<%RsUnitOrgSAP.movefirst
	do until RsUnitOrgSAP.eof
		if PurchasingOrg = trim(RsUnitOrgSAP("PurchasingOrg")) then %>
			<OPTION Value ="<%=trim(RsUnitOrgSAP("PurchasingOrg"))%>" selected >
			<%=RsUnitOrgSAP("PurchasingOrg")%></OPTION>
		<%else%>
			<OPTION Value ="<%=trim(RsUnitOrgSAP("PurchasingOrg"))%>">
			<%=RsUnitOrgSAP("PurchasingOrg")%></OPTION>
		<%end if
		RsUnitOrgSAP.movenext
	loop%>
	</SELECT></td>
	<tr>
	<td height="24" nowrap> Securisation Date :<td width="14%" height="24">
	<input type="text" value="<%=SecureDate%>" size="14"name="SecureDate" style="border-style: inset; border-width: 2px"></td>
	<td width="15%"></td><td  onMouseover="change(this)" onMouseout="change_back(this)" border="1" style="border: 1px dashed #CCCCFF">
	<a href="../allez.htm">Show Delivery Schedule</a></td>
		</tr>
	<tr><!-- Ligne suivante ---------------------------------->
	<td height="24" nowrap> <input type="reset" value="Cancel" name="B2" style="color: #00CC00"><td width="14%" height="24">
	<input type="submit" value="Save" name="BSave" style="color: #0000FF"></td>
	<td width="15%"></td>
	<td  onMouseover="change(this)" onMouseout="change_back(this)" border="1" style="border: 1px dashed #CCCCFF"><a href="../allez.htm">Show Indices</a></td>
	<td colspan="2">
	<input type="submit" value="Delete" name="BtDelete" 
	onmousedown="if (confirm('Warning! \n \n YOU ARE GOING TO DELELTE PERMANENTLY THIS BUSINESS CASE! \n \n Click on \'Cancel\' to Abort')==true) {this.click()};" style="color: #FF0000; font-weight: bold; float: right"></td>
	</tr> 
	</table>
	
	<table><tr><td height="8" nowrap></td></tr></table>
	
	</tr>
	
	<table border="0" width="100%" id="table1" style="border-style: outset; border-width: 3px; background-color: #CCCCFF">
	<tr>
		<%if onglet = "Base line" or (onglet = "" and Status = "B") then %>
		<td width="113" align="center" bgcolor="#9999FF"> 
		<input type="submit" value="Base line" name="onglet" style="color: #0000FF; font-weight: bold">
		<%else%>
		<td width="113" align="center"> 
		<input type="submit" value="Base line" name="onglet">
		<%end if
response.write "</td>"
if onglet = "Secured" or onglet = "Lock and Save" or _
		(onglet = "" and Status = "S") then %>
			<td width="114" align="center" bgcolor="#9999FF">
			<input type="submit" value="Secured" name="onglet" style="color: #0000FF; font-weight: bold">
		<%elseif Status = "B" then%>
			<td width="114" align="center">
			<input type="submit" disabled value="Secured" name="onglet">
		<%else%>
			<td width="114" align="center">
			<input type="submit" value="Secured" name="onglet">
		<%end if%></td>
		<td width="114" align="center">
		<% if Status = "S" then %>
		<input type="button" value="Realised Savings" name="onglet" style="padding: 0"
		onclick="OuvrirFenetre ('RealisedSavingsCalc_ASRaM.asp?status=S','Realised Savings Calculation','status')" >
		<%else%>
			<input type="submit" disabled value="Realised Savings" name="onglet">
		<%end if%></td>
		<td width="114" align="center">
		<% if Status = "S" then %>
		<input type="button" value="Report" name="onglet" style="padding: 0"
		onclick="OuvrirFenetre ('Report_ASRaM_Detail.asp?status=S','Report','target=\'mainFrame\'')" >
		<%else%>
			<input type="submit" disabled value="Report" name="onglet">
		<%end if
		response.write "</td><td>"
 		if Status = "S" then %>
		<input type="button" value="GAP Analysis" name="onglet" style="padding: 0"
		onclick="OuvrirFenetre ('Texte_ASRaM.asp?','GAP Analysis','target=\'mainFrame\'')" >
		<%else%>
			<input type="submit" disabled value="GAP Analysis" name="onglet">
		<%end if
		response.write "</td>"%>
	</tr>
	</table>

<%if onglet = "Base line" or (onglet = "" and Status = "B") then '_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
	set RsEscal =connection.execute("select * from Adv_Escalation where IdBusCase = '" & _
	IdBusCase & "' and Status = 'B'")
 	if not (RsEscal.eof and RsEscal.bof) then ' Fmul existe
 		RsEscal.movefirst
		EscalationFm = trim(RsEscal("EscalationFm"))
	end if
	set RsArticle =connection.Execute("select * from Adv_Article, Adv_Index_BC_Article where Adv_Article.IdArticle=Adv_Index_BC_Article.IdArticle AND IdBusCase = '" & IdBusCase & "' AND Status = 'B' order by TabInd, IdLine, ValidTo desc")

	%>
		
	<table border="1" style="border-collapse: collapse; border: 3px ridge #FFCCFF" width="100%">
	<tr width="100%"><!-- Ligne suivante ---------------------------------->
		<td height="24" style="border-style: ridge; border-width: 2px; padding: 0" bgcolor="#DCE9F6" bordercolordark="#FF0000" width="266" bordercolorlight="#FF0000"> 
			Base line Escalation Formula : 
			<%if Status = "B" then %>
			<img onclick= "OuvrirFenetre ('EscalationFm_ASRaM.asp?status=B','Baseline Formula','fullscreen');" border="0" id="img01" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img01',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img01',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img01',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img0',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit">
			<%end if%>
		</td>
		<td height="24" style="border:2px solid #FF0000; padding:1px; " bordercolorlight="#FF0000" bordercolordark="#000000" bordercolor="#FF0000">
		<%=Replace(EscalationFm,";"," ")%>
		</td>
		<td nowrap style="border-style: ridge; border-width: 2px; padding: 0" bgcolor="#DCE9F6" width="140" bordercolor="#C0C0C0"> 
		Contract Base Year : </td>
		<td style="border-style: ridge; border-width: 2px; padding: 0" bgcolor="#FFFFFF" bordercolordark="#C0C0C0" width="81">
		<%If not RsEscal.eof then response.write RsEscal("iYear")%>
		</td>
	</tr>	
	</table>
	<table border="1" style="border-collapse: collapse; border: 3px ridge #FFCCFF" width="100%">
		<tr>
		<td width="25%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		
			<%
			iTab = 0
			iLine = -1
			i=0
			Total_Art_Price=0
			
			if RsArticle.eof then 'vide
			%>
			<table border="2" width="724">
				<tr>
				<td width="440">
				
				<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF">
				<tr><!-- Ligne suivante ---------------------------------->
					<td height="24" style="border-style: ridge; border-width: 0px; padding: 0" colspan="2" bordercolor="#DCE9F6"> 
					Part numbers 
				 	</td>
					<td style="border-style: ridge; border-width: 0px;" colspan="2">
					<%if Status = "B" then %>
					<img onclick= "OuvrirFenetre('Articles_ASRaM.asp?Status=B&TabInd='+<%=1%>+'&IdCurrency='+document.getElementById('IdCurrency').value)" border="0" id="img02" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img02',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img02',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img02',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img02',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit">
					<%end if%>
				 	</td></td>
					<td style="border-style: ridge; border-width: 0px;">Currency:</td>
					<td style="border-style: ridge; border-width: 0px;" colspan="2">
					<input type="text" readonly name="IdCurrency"  size="14" value ="Eu" style="border-style: solid; border-width: 0px">
				 	</td>
				</tr>
				<tr>
					<td width="2" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					L</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" nowrap>
					Article Code</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Change</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Designation</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" nowrap>
					Valid To</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Price</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" width="9%">
					Qty</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Unit</td>
				</tr>	
			</table>
			<%end if
			
			do until RsArticle.eof
			
	if iTab <> RsArticle("TabInd") then '1ère ligne 
	%>
			<table border="2" width="724">
				<tr>
				<td width="440">
				
				<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF" id="table4">
				<tr><!-- Ligne suivante ---------------------------------->
					<td height="24" style="border-style: ridge; border-width: 0px; padding: 0" colspan="2" bordercolor="#DCE9F6"> 
					Part numbers 
				 	</td>
					<td style="border-style: ridge; border-width: 0px;" colspan="2">
					<%if Status = "B" then %>
					<img onclick= "OuvrirFenetre('Articles_ASRaM.asp?Status=B&TabInd='+<%=RsArticle("TabInd")%>+'&IdCurrency='+document.getElementById('IdCurrency').value)" border="0" id="img02" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img02',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img02',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img02',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img02',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit"> 
					<%end if%>
					</td></td>
					<td style="border-style: ridge; border-width: 0px;">Currency:</td>
					<td style="border-style: ridge; border-width: 0px;" colspan="2">
					<%	RsCurrency.movefirst
						do until RsCurrency.eof
							if trim(RsArticle("IdCurrency")) = trim(RsCurrency("IdCurrency")) then
								response.write trim(RsCurrency("Designation")) & _
							"<input type=""hidden"" name=""IdCurrency"" size=""14"" value =""" & RsArticle("IdCurrency")&""">"
							end if
							RsCurrency.movenext
						loop%>
				 	</td>
				</tr>
				<tr>
					<td width="2" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					L</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" nowrap>
					Article Code</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Change</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Designation</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" nowrap>
					Valid To</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Price</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" width="9%">
					Qty</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Unit</td>
				</tr>	
				<tr>
				<% if RsArticle("ValidTo") < 9999 then 'Si l'Article n'est plus valide%>
					<td width="2" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<font color="#C0C0C0"><%=RsArticle("IdLine")%></font></td>
					<td width="15%" nowrap class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<font color="#C0C0C0"><%=RsArticle("IdArticle")%></font></td>
					<td style="font-family: Wingdings" align="center" bordercolor="#000000"><b>
					<font face="Webdings" color="#D6A3A3">x</font></b></td>
					<td nowrap class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Designation")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<font color="#C0C0C0"><%=RsArticle("ValidTo")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Art_Price")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Art_Qty")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
				<%	do until RsUnit.eof
						if RsArticle("IdUnit")=RsUnit("IdUnit") then
							Response.write "<font color=""#C0C0C0"">"&RsUnit("Designation")
						end if
						RsUnit.movenext
					loop
					RsUnit.movefirst%>
					</td>
				<%else 'si Article valide
				i=i+1%>
					<td width="2%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<%=RsArticle("IdLine")%></td>
					<td width="22%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<%=RsArticle("IdArticle")%></td>
					<td align="center" bordercolor="#000000">
					<%if Status = "B" Then %>
						<IMG onclick= "OuvrirFenetre('Article_Replace_ASRaM.asp?IdLine=<%=trim(RsArticle("IdLine"))%>&TabInd=<%=trim(RsArticle("TabInd"))%>&IdUnit=<%=trim(RsArticle("IdUnit"))%>&IdArticle=<%=trim(RsArticle("IdArticle"))%>&Art_Qty=<%=trim(RsArticle("Art_Qty"))%>&Art_Price=<%=trim(RsArticle("Art_Price"))%>&Status=<%=trim(RsArticle("Status"))%>&ValidTo=<%=trim(RsArticle("ValidTo"))%>&Designation=<%=trim(RsArticle("Designation"))%>&IdCurrency='+document.getElementById('IdCurrency').value)" border=0 height=18 hspace=5 src="images/Create.JPG" width=17>
					<%end if%>
					</td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center" width="33%"><%=RsArticle("Designation")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center" width="33%"><%=RsArticle("ValidTo")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center"><%=RsArticle("Art_Price")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center"><%=RsArticle("Art_Qty")%></td>
					
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<%if trim(RsArticle("ValidTo"))="9999" then Total_Art_Price=Total_Art_Price+RsArticle("Art_Price")
					do until RsUnit.eof
					if RsArticle("IdUnit")=RsUnit("IdUnit") then
						response.write RsUnit("Designation")
					end if
					RsUnit.movenext
					loop
					RsUnit.movefirst%>
					</td>
					<%				end if

		else '2ème ligne jusqu'à fin
				%>
				
				<tr>
				<% if RsArticle("ValidTo") < 9999 then 'Si la ligne existe déjà %>
					<td width="2" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<font color="#C0C0C0"><%=RsArticle("IdLine")%></font></td>
					<td width="15%" nowrap class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<font color="#C0C0C0"><%=RsArticle("IdArticle")%></font></td>
					<td align="center" bordercolor="#000000"><b><font face="Webdings" color="#D6A3A3">
					x</font></b></td>
					<td nowrap class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Designation")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<font color="#C0C0C0"><%=RsArticle("ValidTo")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Art_Price")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Art_Qty")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<%do until RsUnit.eof
					if RsArticle("IdUnit")=RsUnit("IdUnit") then
						Response.write "<font color=""#C0C0C0"">" & RsUnit("Designation")
					end if
					RsUnit.movenext
					loop
					RsUnit.movefirst%>
					</td>

					<%else%>
					<td width="2%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<%=RsArticle("IdLine")%></td>
					<td width="22%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<%=RsArticle("IdArticle")%></td>
					<td align="center" bordercolor="#000000">
					<%if Status = "B" Then %>
						<IMG onclick= "OuvrirFenetre('Article_Replace_ASRaM.asp?IdLine=<%=trim(RsArticle("IdLine"))%>&TabInd=<%=trim(RsArticle("TabInd"))%>&IdUnit=<%=trim(RsArticle("IdUnit"))%>&IdArticle=<%=trim(RsArticle("IdArticle"))%>&Art_Qty=<%=trim(RsArticle("Art_Qty"))%>&Art_Price=<%=trim(RsArticle("Art_Price"))%>&Status=<%=trim(RsArticle("Status"))%>&ValidTo=<%=trim(RsArticle("ValidTo"))%>&Designation=<%=trim(RsArticle("Designation"))%>&IdCurrency='+document.getElementById('IdCurrency').value)" border=0 height=18 hspace=5 src="images/Create.JPG" width=17>
					<%end if%>
					</td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center" width="33%"><%=RsArticle("Designation")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center" width="33%"><%=RsArticle("ValidTo")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center"><%=RsArticle("Art_Price")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center"><%=RsArticle("Art_Qty")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<%if trim(RsArticle("ValidTo"))="9999" then Total_Art_Price=Total_Art_Price+RsArticle("Art_Price")
					do until RsUnit.eof
					if RsArticle("IdUnit")=RsUnit("IdUnit") then
						Response.write RsUnit("Designation")
					end if
					RsUnit.movenext
					loop
					RsUnit.movefirst%>
					</td>
				<%end if
			end if
			

			iTab = RsArticle("TabInd")
			iLine = RsArticle("IdLine")

			RsArticle.movenext
			if not ( RsArticle.eof ) then
				if iTab <> RsArticle("TabInd") then
				%>
				</tr>				
				</table>				
					<td>
						<table>				
							<tr>
							<td width="23%" height="24" style="border-style: ridge; border-width: 0px; padding: 0" bordercolordark="#C0C0C0"> 
							Variants </td>
							<td width="57%" height="24" style="border-style: ridge; border-width: 0px">			
							<%if Status = "B" then %>
							<img onclick= "OuvrirFenetre('Variants_ASRaM.asp?Status=B&TabInd='+<%=iTab%>)" border="0" id="img03" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img03',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img03',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img03',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img03',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit">
							<%end if%>
							</td>
							</tr>
						</table>				
						<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF" id="table4">
							<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">IdVariant</td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%">Designation</td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%">Var_Qty</td>
					<%
					set RsVariant =connection.Execute("select * from Adv_Variant, Adv_Index_BC_Variant where Adv_Variant.IdVariant=Adv_Index_BC_Variant.IdVariant AND IdBusCase = '"& IdBusCase & "' AND TabInd = "& iTab & " AND Status = 'B'")
					if not (RsVariant.eof and RsVariant.bof) then
					IdMacro=RsVariant("IdMacro")
					do until RsVariant.eof %>
							<tr>
							<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF"><%=RsVariant("IdVariant")%></td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%"><%=RsVariant("Designation")%></td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%"><%=RsVariant("Var_Qty")%></td>
							</tr>
				<% RsVariant.movenext
				 	Loop
					end if				
					Response.write "</table>" & _ 
						"IdMacro <input type=""text"" readonly name=""IdMacro"" style=""border: 1px solid #000000"" size=""14"" value="&IdMacro&" >" & _
						"</td>" & _
					"</tr>"& _
				"</table><font color=""#800080"">Total __________________________________________  "& _
					Total_Art_Price 
					else
					Response.write "</tr>"
					end if
				else				
				%>
				</tr>
				</table>
						<td>
						<table>				
							<td class="titrelog" style="border: 0px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" width="2%">
							</td>
							<td width="23%" height="24" style="border-style: ridge; border-width: 0px; padding: 0" bordercolordark="#C0C0C0"> 
							Variants </td>
							<td width="57%" height="24" style="border-style: ridge; border-width: 0px">
							<%if Status = "B" then %>
							<img onclick= "javascript:OuvrirFenetre('Variants_ASRaM.asp?Status=B&TabInd='+<%=iTab%>)" border="0" id="img04" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img04',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img04',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img04',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img04',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit">
							<%end if%>
							</td>
						</table>				
						<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF" id="table4">
							<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">IdVariant</td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%">Designation</td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%">Var_Qty</td>
					<%
					set RsVariant =connection.Execute("select * from Adv_Variant, Adv_Index_BC_Variant where Adv_Variant.IdVariant=Adv_Index_BC_Variant.IdVariant AND IdBusCase = '"& IdBusCase & "' AND TabInd = "& iTab & " AND Status = 'B'")
					if not (RsVariant.eof and RsVariant.bof) then
					IdMacro=RsVariant("IdMacro")
					do until RsVariant.eof %>
							<tr>
							<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF"><%=RsVariant("IdVariant")%></td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%"><%=RsVariant("Designation")%></td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%"><%=RsVariant("Var_Qty")%></td>
							</tr>
				<% RsVariant.movenext
				 	Loop
					end if
					Response.write "</table>" & _
						"IdMacro <input type=""text"" readonly name=""IdMacro"" style=""border: 1px solid #000000"" size=""14"" value="&IdMacro&" >" & _
						"</td>" & _
					"</tr>"& _
				"</table><font color=""#800080"">Total __________________________________________  "& _
					Total_Art_Price 

			Total_Art_Price=0
			
			end if
			loop
		
	if Status = "B" then %>
	<table border="0" width="100%" id="table5" style="border-style: outset; border-width: 3px; ">
	<tr>
		<td width="221" align="center" nowrap>
		<input type="button" onclick= "OuvrirFenetre('Articles_ASRaM.asp?Status='+'B'+'&TabInd='+<%=iTab+1%>+'&IdCurrency=Eu')" value="New Macro" name="NewMacro" style="color: #00CC00">
		</td>
		<td align="right">
		<%If EscalationFm <>"" then %>
		<input type="submit" value="Lock and Save" name="onglet" style="color: #FF0000"
		 onmousedown="if (confirm('Warning! \n If you confirm, the baseline will be locked')==true) {this.click()};" >
		<%else%>
		<input disabled type="submit" value="Lock and Save" name="onglet" style="color: #FF0000">
		<%end if%>
		</td>
	</tr>
	</table>
	<%end if%>
	</table>
	
	
<%'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
elseif onglet = "Secured" or onglet = "Lock and Save" or _
(onglet = "" and Status = "S") then' _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_

	set RsEscal =connection.execute("select * from Adv_Escalation where IdBusCase = '" & _
	 IdBusCase & "' and Status = 'S'")
 	if not (RsEscal.eof and RsEscal.bof) then ' Fmul existe
 		RsEscal.movefirst
		EscalationFm = trim(RsEscal("EscalationFm"))
	end if
	set RsArticle =connection.Execute("select * from Adv_Article, Adv_Index_BC_Article where Adv_Article.IdArticle=Adv_Index_BC_Article.IdArticle AND IdBusCase = '" & IdBusCase & "' AND Status = 'S' order by TabInd, IdLine, ValidTo desc")
	%>
		
	<table border="1" style="border-collapse: collapse; border: 3px ridge #FFCCFF" width="100%">
	<tr width="100%"><!-- Ligne suivante ---------------------------------->
		<td height="24" style="border-style: ridge; border-width: 2px; padding: 0" bgcolor="#DCE9F6" bordercolordark="#FF0000" width="263" bordercolorlight="#FF0000"> 
		Secured Escalation Formula : 
		<%if Status = "S" then %>
			<img onclick= "OuvrirFenetre ('EscalationFm_ASRaM.asp?status=S','Secured Formula','fullscreen');" border="0" id="img01" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img01',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img01',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img01',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img0',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit">
		<%end if%>
		</td>
		<td height="24" style="border-style: ridge; border-width: 2px; padding: 1px" bordercolorlight="#FF0000" bordercolordark="#FF0000">
		<%=Replace(EscalationFm,";"," ")%>
		</td>
		<td style="border-style: ridge; border-width: 2px; padding: 0" bgcolor="#DCE9F6" width="140" bordercolor="#C0C0C0"> 
		Contract Base Year : </td>
		<td style="border-style: ridge; border-width: 2px; padding: 0" bgcolor="#FFFFFF" bordercolordark="#C0C0C0" width="90">
		<%If not RsEscal.eof then response.write RsEscal("iYear")%>
		</td>
	</tr>	
	</table>
	<table border="1" style="border-collapse: collapse; border: 3px ridge #FFCCFF" width="100%">
		<tr>
		<td width="25%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
		
			<%
			iTab = 0
			iLine = -1
			i=0
			
			if RsArticle.eof then 'vide
			%>
			<table border="2" width="724">
				<tr>
				<td width="440">
				
				<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF" id="table4">
				<tr><!-- Ligne suivante ---------------------------------->
					<td height="24" style="border-style: ridge; border-width: 0px; padding: 0" colspan="2" bordercolor="#DCE9F6"> 
					Part numbers 
				 	</td>
					<td style="border-style: ridge; border-width: 0px;" colspan="2">
					<img onclick= "OuvrirFenetre('Articles_ASRaM.asp?Status=S&TabInd='+<%=1%>+'&IdCurrency='+document.getElementById('IdCurrency').value)" border="0" id="img02" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img02',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img02',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img02',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img02',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit"> </td>
				 	</td>
					<td style="border-style: ridge; border-width: 0px;">Currency:</td>
					<td style="border-style: ridge; border-width: 0px;" colspan="2">
					<input type="text" readonly name="IdCurrency"  size="14" value ="Eu" style="border-style: solid; border-width: 0px">
				 	</td>
				</tr>
				<tr>
					<td width="2" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					L</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" nowrap>
					Article Code</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Change</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Designation</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" nowrap>
					Valid To</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Price</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" width="9%">
					Qty</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Unit</td>
				</tr>	
			</table>
			<%end if
			
			do until RsArticle.eof
			
	if iTab <> RsArticle("TabInd") then '1ère ligne 
	%>
			<table border="2" width="724">
				<tr>
				<td width="440">
				
				<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF" id="table4">
				<tr><!-- Ligne suivante ---------------------------------->
					<td height="24" style="border-style: ridge; border-width: 0px; padding: 0" colspan="2" bordercolor="#DCE9F6"> 
					Part numbers 
				 	</td>
					<td style="border-style: ridge; border-width: 0px;" colspan="2">
					<img onclick= "OuvrirFenetre('Articles_ASRaM.asp?Status=S&TabInd='+<%=RsArticle("TabInd")%>+'&IdCurrency='+document.getElementById('IdCurrency').value)" border="0" id="img02" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img02',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img02',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img02',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img02',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit"> </td>
				 	</td>
					<td style="border-style: ridge; border-width: 0px;">Currency:</td>
					<td style="border-style: ridge; border-width: 0px;" colspan="2">
					<%RsCurrency.movefirst
					do until RsCurrency.eof
						if trim(RsArticle("IdCurrency")) = trim(RsCurrency("IdCurrency")) then
							response.write trim(RsCurrency("Designation"))%>
						<input type="hidden" name="IdCurrency" size="13" value ="<%=RsArticle("IdCurrency")%>">
						<%end if
						RsCurrency.movenext
					loop%>
					</SELECT>
				 	</td>
				</tr>
				<tr>
					<td width="2" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					L</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" nowrap>
					Article Code</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Change</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Designation</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" nowrap>
					Valid To</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Price</td>
					<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" width="9%">
					Qty</td>
					<td width="33%" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
					Unit</td>
				</tr>	
				<tr>
				<% if RsArticle("ValidTo") < 9999 then 'Si l'Article n'est plus valide%>
					<td width="2" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<font color="#C0C0C0"><%=RsArticle("IdLine")%></font></td>
					<td nowrap width="15%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<font color="#C0C0C0"><%=RsArticle("IdArticle")%></font></td>
					<td style="font-family: Wingdings" align="center" bordercolor="#000000"><b>
					<font face="Webdings" color="#D6A3A3">x</font></b></td>
					<td nowrap class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Designation")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<font color="#C0C0C0"><%=RsArticle("ValidTo")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Art_Price")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Art_Qty")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<%do until RsUnit.eof
					if RsArticle("IdUnit")=RsUnit("IdUnit") then
						Response.write "<font color=""#C0C0C0"">" & RsUnit("Designation")
					end if
					RsUnit.movenext
					loop
					RsUnit.movefirst%>
					</td>
					<%				else 'si Article valide
				i=i+1%>
					<td width="2%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<%=RsArticle("IdLine")%></td>
					<td width="22%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<%=RsArticle("IdArticle")%></td>
					<td align="center" bordercolor="#000000">
					<%if Status = "S" Then %>
						<IMG onclick= "OuvrirFenetre('Article_Replace_ASRaM.asp?IdLine=<%=trim(RsArticle("IdLine"))%>&TabInd=<%=trim(RsArticle("TabInd"))%>&IdUnit=<%=trim(RsArticle("IdUnit"))%>&IdArticle=<%=trim(RsArticle("IdArticle"))%>&Art_Qty=<%=trim(RsArticle("Art_Qty"))%>&Art_Price=<%=trim(RsArticle("Art_Price"))%>&Status=<%=trim(RsArticle("Status"))%>&ValidTo=<%=trim(RsArticle("ValidTo"))%>&Designation=<%=trim(RsArticle("Designation"))%>&IdCurrency='+document.getElementById('IdCurrency').value)" border=0 height=18 hspace=5 src="images/Create.JPG" width=17>
					<%end if%>
					</td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center" width="33%"><%=RsArticle("Designation")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center" width="33%"><%=RsArticle("ValidTo")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center"><%=RsArticle("Art_Price")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center"><%=RsArticle("Art_Qty")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<%if trim(RsArticle("ValidTo"))="9999" then Total_Art_Price=Total_Art_Price+RsArticle("Art_Price")
					do until RsUnit.eof
					if RsArticle("IdUnit")=RsUnit("IdUnit") then
						response.write RsUnit("Designation")
					end if
					RsUnit.movenext
					loop
					RsUnit.movefirst%>
					</td>
					<%				end if

		else '2ème ligne jusqu'à fin
				%>
				
				<tr>
				<% if RsArticle("ValidTo") < 9999 then 'Si la ligne existe déjà %>
					<td width="2" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<font color="#C0C0C0"><%=RsArticle("IdLine")%></font></td>
					<td nowrap width="15%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<font color="#C0C0C0"><%=RsArticle("IdArticle")%></font></td>
					<td align="center" bordercolor="#000000"><b><font face="Webdings" color="#D6A3A3">
					x</font></b></td>
					<td nowrap class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Designation")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<font color="#C0C0C0"><%=RsArticle("ValidTo")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Art_Price")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center">
					<font color="#C0C0C0"><%=RsArticle("Art_Qty")%></font></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<%do until RsUnit.eof
					if RsArticle("IdUnit")=RsUnit("IdUnit") then
						Response.write "<font color=""#C0C0C0"">" & RsUnit("Designation")
					end if
					RsUnit.movenext
					loop
					RsUnit.movefirst%>
					</td>
					<%					else%>
					<td width="2%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<%=RsArticle("IdLine")%></td>
					<td width="22%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center">
					<%=RsArticle("IdArticle")%></td>
					<td align="center" bordercolor="#000000">
					<%if Status = "S" Then %>
						<IMG onclick= "OuvrirFenetre('Article_Replace_ASRaM.asp?IdLine=<%=trim(RsArticle("IdLine"))%>&TabInd=<%=trim(RsArticle("TabInd"))%>&IdUnit=<%=trim(RsArticle("IdUnit"))%>&IdArticle=<%=trim(RsArticle("IdArticle"))%>&Art_Qty=<%=trim(RsArticle("Art_Qty"))%>&Art_Price=<%=trim(RsArticle("Art_Price"))%>&Status=<%=trim(RsArticle("Status"))%>&ValidTo=<%=trim(RsArticle("ValidTo"))%>&Designation=<%=trim(RsArticle("Designation"))%>&IdCurrency='+document.getElementById('IdCurrency').value)" border=0 height=18 hspace=5 src="images/Create.JPG" width=17>
					<%end if%>
					</td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center" width="33%"><%=RsArticle("Designation")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" align="center" width="33%"><%=RsArticle("ValidTo")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center"><%=RsArticle("Art_Price")%></td>
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%" align="center"><%=RsArticle("Art_Qty")%></td>
					
					<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="33%" align="center">
					<%if trim(RsArticle("ValidTo"))="9999" then Total_Art_Price=Total_Art_Price+RsArticle("Art_Price")
					do until RsUnit.eof
					if RsArticle("IdUnit")=RsUnit("IdUnit") then
						Response.write RsUnit("Designation")
					end if
					RsUnit.movenext
					loop
					RsUnit.movefirst%>
					</td>
					<%	end if%>
		<%end if	
			
			iTab = RsArticle("TabInd")
			iLine = RsArticle("IdLine")
			RsArticle.movenext
			if not ( RsArticle.eof ) then
				if iTab <> RsArticle("TabInd") then
				%>
				</tr>
				</table>				
					<td>
						<table>				
							<tr>
							<td width="23%" height="24" style="border-style: ridge; border-width: 0px; padding: 0" bordercolordark="#C0C0C0"> 
							Variants </td>
							<td width="57%" height="24" style="border-style: ridge; border-width: 0px">			
							<img onclick= "OuvrirFenetre('Variants_ASRaM.asp?Status=S&TabInd='+<%=iTab%>)" border="0" id="img03" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img03',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img03',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img03',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img03',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit">
							</td>
							</tr>
						</table>				
						<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF" id="table4">
							<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">IdVariant</td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%">Designation</td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%">Var_Qty</td>
					<%
					set RsVariant =connection.Execute("select * from Adv_Variant, Adv_Index_BC_Variant where Adv_Variant.IdVariant=Adv_Index_BC_Variant.IdVariant AND IdBusCase = '"& IdBusCase & "' AND TabInd = "& iTab & " AND Status = 'S'")
					if not (RsVariant.eof and RsVariant.bof) then
						IdMacro=RsVariant("IdMacro")
						do until RsVariant.eof %>
							<tr>
							<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF"><%=RsVariant("IdVariant")%></td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%"><%=RsVariant("Designation")%></td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%"><%=RsVariant("Var_Qty")%></td>
							</tr>
					<% RsVariant.movenext
				 	Loop
					end if				
					Response.write "</table>" & _ 
						"IdMacro <input type=""text"" readonly name=""IdMacro"" style=""border: 1px solid #000000"" size=""14"" value="&IdMacro&" >" & _
						"</td>" & _
					"</tr>"& _
				"</table><font color=""#800080"">Total __________________________________________  "& _
					Total_Art_Price 
					else
					Response.write "</tr>"
					end if
				else				
				%>
				</tr>
				</table>
						<td>
						<table height="26">				
							<td class="titrelog" style="border: 0px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6" width="2%">
							</td>
							<td width="23%" height="22" style="border-style: ridge; border-width: 0px; padding: 0" bordercolordark="#C0C0C0"> 
							Variants </td>
							<td width="57%" height="22" style="border-style: ridge; border-width: 0px">			
							<img onclick= "javascript:OuvrirFenetre('Variants_ASRaM.asp?Status=S&TabInd='+<%=iTab%>)" border="0" id="img04" src="images/buttonEdit.jpg" height="15" width="77" alt=" Edit" onmouseover="FP_swapImg(1,0,/*id*/'img04',/*url*/'images/button1.jpg')" onmouseout="FP_swapImg(0,0,/*id*/'img04',/*url*/'images/buttonEdit.jpg')" onmousedown="FP_swapImg(1,0,/*id*/'img04',/*url*/'images/button2.jpg')" onmouseup="FP_swapImg(0,0,/*id*/'img04',/*url*/'images/button1.jpg')" fp-style="fp-btn: Jewel 1; fp-justify-horiz: 0" fp-title=" Edit">
							</td>
						</table>				
						<table border="1"  width="100%" style="border-collapse: collapse; border: 3px ridge #FFCCFF" id="table4">
							<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">IdVariant</td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%">Designation</td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%">Var_Qty</td>
					<%
					set RsVariant =connection.Execute("select * from Adv_Variant, Adv_Index_BC_Variant where Adv_Variant.IdVariant=Adv_Index_BC_Variant.IdVariant AND IdBusCase = '"& IdBusCase & "' AND TabInd = "& iTab & " AND Status = 'S'")
					if not (RsVariant.eof and RsVariant.bof) then
						IdMacro=RsVariant("IdMacro")
					do until RsVariant.eof %>
							<tr>
							<td width="25%" class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF"><%=RsVariant("IdVariant")%></td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="6%"><%=RsVariant("Designation")%></td>
							<td class="rublog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF" width="9%"><%=RsVariant("Var_Qty")%></td>
							</tr>
				<% RsVariant.movenext
				 	Loop
					end if				
					Response.write "</table>" & _ 
						"IdMacro <input type=""text"" readonly name=""IdMacro"" style=""border: 1px solid #000000"" size=""14"" value="&IdMacro&" >" & _
						"</td>" & _
					"</tr>"& _
				"</table><font color=""#800080"">Total __________________________________________  "& _
					Total_Art_Price 

			end if
			loop
		%>
		
	<table border="0" width="100%" id="table5" style="border-style: outset; border-width: 3px; ">
	<tr>
		<td width="221" align="center" nowrap>
		<% RsArticle.movefirst
		if not RsArticle.eof and RsVariant.eof and RsVariant.bof then%>
		<input type="button" disabled value="New Macro" name="NewMacro" style="color: #00CC00">
		<%else%>
		<input type="button" onclick= "OuvrirFenetre('Articles_ASRaM.asp?Status='+'S'+'&TabInd='+<%=iTab+1%>+'&IdCurrency=Eu')" value="New Macro" name="NewMacro" style="color: #00CC00">
		<%end if%>
		</td>
	</tr>
	</table>
	</table>


<%' _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
elseif onglet = "Realised Savings" then '_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
%>

<% elseif onglet = "Report" then ' _-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_ %>
	<input type="text" value="<%=onglet%>" size="14"name="Test">
<%end if%>
</form>

</body>
<script language="JavaScript">
<!--#INCLUDE FILE="connection/deconnection.asp"-->
</script>
</html>