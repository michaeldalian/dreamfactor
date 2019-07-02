<html>
	<head>
		<script language="javascript">
			<!--#INCLUDE FILE="connection/connection.asp"-->
			function increment_year()
			{
				document.debut.annee0.value ++
				//verifier()
			}
			function decrement_year()
			{
				document.debut.annee0.value=document.debut.annee0.value -1
				//verifier()
			}
			function copie()
			{
				document.debut.id_indice.value=document.debut.id_indice0.value
				document.debut.annee.value=document.debut.annee0.value
			}
			function CtrlVal(Nom)
			{
				// Procedure de validation
				if(window.event.keyCode==46 || window.event.keyCode==44)
				{
					window.event.keyCode = 46;
				}				
				if (window.event.keyCode<48 || window.event.keyCode>57)
				{
					if(window.event.keyCode=46)
					{
						window.event.keyCode=46;// On autorise la frappe de la touche 46 qui est inférieur à 48 qui est en fait le .
					}
					else
					window.event.keyCode = 0;
				}
			}
			function verifier()
			{
				var validation=true
				if(document.debut.value1.value=="" && document.debut.value2.value=="" && document.indice_update.value3.value=="" && document.debut.value4.value=="" && document.debut.value5.value=="" && document.debut.value6.value=="" && document.debut.value7.value=="" && document.debut.value8.value=="" && document.debut.value9.value=="" && document.debut.value10.value=="" && document.debut.value11.value=="" && document.debut.value12.value=="")
				{
					alert("All values can't be null")
					validation=false
				}
				//var maReg = new RegExp ( /[0-9]{0,1}[,]{0,1}[0-9]/ ) ;
				var maReg = new RegExp (/\d+\,?\d*|\,\d+/);
				if (document.debut.value1.value!='' && document.debut.value1.value.search(maReg ) == -1  && document.debut.value1.value!=0)
					{
						alert ( "The January value is not at the good format" ) ;
						validation=false
					}
						
					if (document.debut.value2.value!='' && document.debut.value2.value.search(maReg ) == -1  && document.debut.value2.value!=0)
					{

						alert ( "The February value is not at the good format" ) ;
						validation=false
					}
					if (document.debut.value3.value!='' && document.debut.value3.value.search(maReg ) == -1  && document.debut.value3.value!=0)
					{
						alert ( "The March value is not at the good format" ) ;
						validation=false
					}
					if (document.debut.value4.value!='' && document.debut.value4.value.search(maReg ) == -1  && document.debut.value4.value!=0)
					{
						alert ( "The April value is not at the good format" ) ;
						validation=false

					}
					if (document.debut.value5.value!='' && document.debut.value5.value.search(maReg ) == -1  && document.debut.value5.value!=0)
					{

						alert ( "The May value is not at the good format" ) ;
						validation=false
					}
					if (document.debut.value6.value!='' && document.debut.value6.value.search(maReg ) == -1  && document.debut.value6.value!=0)
					{
						alert ( "The June value is not at the good format" ) ;
						validation=false
					}
					if (document.debut.value7.value!='' && document.debut.value7.value.search(maReg ) == -1  && document.debut.value7.value!=0)
					{
						alert ( "The July value is not at the good format" ) ;
						validation=false
					}
					if (document.debut.value8.value!='' && document.debut.value8.value.search(maReg ) == -1  && document.debut.value8.value!=0)
					{

						alert ( "The August value is not at the good format" ) ;
						validation=false
					}
					if (document.debut.value9.value!='' && document.debut.value9.value.search(maReg ) == -1  && document.debut.value9.value!=0)
					{

						alert ( "The September value is not at the good format" ) ;
						validation=false

					}
					if (document.debut.value10.value!='' && document.debut.value10.value.search(maReg ) == -1  && document.debut.value10.value!=0)
					{

						alert ( "The October value is not at the good format" ) ;
						validation=false
					}
					if (document.debut.value11.value!='' && document.debut.value11.value.search(maReg ) == -1  && document.debut.value11.value!=0)
					{

						alert ( "The November value is not at the good format" ) ;
						validation=false
					}
					if (document.debut.value12.value!='' && document.debut.value12.value.search(maReg ) == -1  && document.debut.value12.value!=0)
					{
						alert ( "The December value is not at the good format" ) ;
						validation=false
					}
					
					//document.indice_update.annee.value=document.debut.annee.value
					//document.indice_update.id_indice.value=document.debut.id_indice.value
				
				if(validation ==true)
				{
					document.debut.submit()
				}
					
				
			}
			OnKeyUp_Verif=function(obj) {
   var virgule=false;
   var oldvalue=obj.value;
   var newvalue=""
   for (var i=0; i<oldvalue.length; i++) {
      if ((oldvalue.charAt(i) == "," || oldvalue.charAt(i) == ".") && !virgule) {
         if (i==0) {
            newvalue+="0"
         }
         newvalue+=","
         virgule=true;
      }
      else if (!isNaN(parseInt("" + oldvalue.charAt(i) + ""))) {
         newvalue+=oldvalue.charAt(i)
      }
   }
   if (oldvalue!=newvalue) {
      obj.value=newvalue;
   }
}
			</script>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<link rel="stylesheet" type="text/css" href="http://intra-dev/intrapur/procurement_mrs/css/log.css">
		<title>Update Indice</title>
		<style>
			all.clsMenuItemNS{font: bold x-small Verdana; color: white; text-decoration: none;}
			.clsMenuItemIE{text-decoration: none; font: bold xx-small Verdana; color: white; cursor: hand;}
			A:hover {color: red;}
		</style>
	</head>
	<body onload="copie();"bgcolor="#FFFFFF" text="#000000" background="http://intra-dev/intrapur/procurement_mrs/images/DefaultHomePage.gif" style="background-repeat: no-repeat;background-attachment: scroll;background-position: center center">
		<p align="center">
		<% 
 		Function DelSpace(Chaine) 
			Dim ChaineCopy, i, NbrCarToEnd, Part1, Part2 
  			ChaineCopy = LTrim(Chaine) 
			ChaineCopy = RTrim(ChaineCopy) 
			i = InStr(1, ChaineCopy, " ") 
			While i <> 0 
    			NbrCarToEnd = Len(ChaineCopy) - (i - 1) 
    			Part1 = Mid(ChaineCopy, 1, i) 
    			Part2 = Mid(ChaineCopy, i + 2, NbrCarToEnd) 
    			ChaineCopy = Part1 & Part2 
    			i = InStr(1, ChaineCopy, " ") 
			Wend 
			DelSpace = ChaineCopy 
		End Function 
		
		'Permet d'afficher tous les parametres reçus
			'For Each chaine_requete In Request.QueryString
  				'Response.Write "<tr>;<td>" & chaine_requete & "</td> <td>= " & Request.QueryString(chaine_requete) & "</td></tr>"       
			'Next
		if (request.querystring("update")=1 ) then
				if not(request.querystring("value1")="" ) then
					ps=" exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/01/01"&  "'," & Replace(request.querystring("value1"),",",".")
					ps=ps & ",'';"
				end if
				if not(request.querystring("value2")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','"& request.querystring("annee") & "/02/01" & "'," & Replace(request.querystring("value2"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value3")="") then
					ps=ps &" exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','"& request.querystring("annee") & "/03/01" & "'," & Replace(request.querystring("value3"),",",".") 
					ps=ps & ",'';"
				end if
				if not(request.querystring("value4")="") then
					ps=ps &" exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','"& request.querystring("annee") & "/04/01" & "'," & Replace(request.querystring("value4"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value5")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/05/01"  & "'," & Replace(request.querystring("value5"),",",".") 
					ps=ps & ",'';"
				end if
				if not(request.querystring("value6")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/06/01"  & "'," & Replace(request.querystring("value6"),",",".") 
					ps=ps & ",'';"
				end if
				if not(request.querystring("value7")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/07/01"  & "'," & Replace(request.querystring("value7"),",",".") 
					ps=ps & ",'';"
				end if
				if not(request.querystring("value8")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/08/01"  & "'," & Replace(request.querystring("value8"),",",".") 
					ps=ps & ",'';"
				end if
				if not(request.querystring("value9")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/09/01"  & "'," & Replace(request.querystring("value9"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value10")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/10/01"  & "'," & Replace(request.querystring("value10"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value11")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/11/01"  & "'," & Replace(request.querystring("value11"),",",".") 
					ps=ps & ",'';"
				end if
				if not(request.querystring("value12")="") then
					ps=ps &"exec ps_adv_MajIndice_Values" & " '" & request.querystring("id_indice") & "','" & request.querystring("annee") & "/12/01"  & "'," & Replace(request.querystring("value12"),",",".")  
					ps=ps & ",'';"
				end if

				response.end
				set update_indice=connection.Execute(ps)
				response.write("Update has been done correctly")
				%>
				<script language="javascript">
					//window.location="ASRAM_update_indice.asp?annee="+ <%=request.querystring("annee")%>
				</script>
				<%
			end if	
		%>
		<p align="center">
			<u><font size="6">
			<table width="465">
				<tr>
					<td align="center">
						<table border="0" width="41%" id="table3">
							<tr>
								<td class="titrelog" align="center">
			<span style="font-weight: 400">
								<font size="5">Indices Update 
						</font></span>
								</td>
							</tr>
						</table>
						<p align="center">&nbsp;<p align="center">&nbsp;
					</td>			
				</tr>
				<tr>
					<td align="center">
						<table border="0" id="table1">
							<tr>
								<td class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
									<form name="debut" action="ASRAM_update_indice.asp" method="get">
									<font size="2">Year</font><font size="4"><b>&nbsp;&nbsp;</td> 	
								<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">
									<img border="0" src="http://intra-dev/intrapur/procurement_mrs/ASRaM/images/ASRAM_1.gif" onclick="decrement_year();">
										<%
										if not(request.querystring("annee0")="") then%>
											<input type="text" name="annee0" value="<%=request.querystring("annee0")%>" length="4" size="4" maxlength="4"  onpropertychange="verifier();"></font></b>
										<%else%>
											<input type="text" name="annee0" value="<%=year(date)%>" length="4" size="4" maxlength="4" onpropertychange="verifier();">
										<%end if%>
									<img border="0" src="http://intra-dev/intrapur/procurement_mrs/ASRaM/images/ASRAM_2.gif" onclick="increment_year();">
								</td>
							</tr>
							<tr>
								<td valign="middle" class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
									<font size="2">Indices</font></td>
							<td valign="middle" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">
								<select name="id_indice0" onchange="verifier();">
									<%	
									set GetIndice=connection.Execute("select * from Adv_Indice order by IdIndice")
									if not GetIndice.eof then
										do until GetIndice.eof
											if(trim(request.querystring("id_indice0"))=trim(GetIndice("IdIndice"))) then
											%>
												<OPTION selected value="<%=trim(GetIndice("IdIndice"))%>"> <%=GetIndice("IdIndice")%> </OPTION>
											<%else%>
												<OPTION value="<%=trim(GetIndice("IdIndice"))%>"> <%=GetIndice("IdIndice")%></OPTION>
											<%
											end if
											GetIndice.movenext
										loop
									end if
									%>
								</select>
								

						</td>
							<input type="hidden" name="update" value=1>
							<input type="hidden" name="id_indice">
							<input type="hidden" name="annee">
					</tr>
					<tr>
						<td valign="top"class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
							<font size="2">Values</font></td>
						<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">
							<table border="0" width="100%" id="table2">
								<tr>
									<td width="79">January</td>
									<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/01/01',102) and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/01/31',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												'response.write(query)
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value1" onblur="verifNombre(this.value);"onkeypress ="javascript:CtrlVal('value1')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value1" value=0 onkeypress ="javascript:CtrlVal('value1')"></td>
												<%
											end if
										else
											set GetIndice=connection.Execute("select top 1 IdIndice from Adv_Indice order by IdIndice")
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/01/01',102) and CONVERT(DATETIME,'" & year(date) & "/01/31',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input onkeyup="OnKeyUp_Verif(this)" type="text" value="<%=getvalueindice("ivalue")%>"name="value1" onkeypress ="javascript:CtrlVal('value1')"></td><%
											else
										
												%>
												<td><input onkeyup="OnKeyUp_Verif(this)" type="text" name="value1" value=0 onchange="check(this.value)"  onkeypress ="javascript:CtrlVal('value1')"></td>
												<%
											end if
										end if	
										%>
								</tr>
								<tr>
									<td width="79">February</td>
										<%
										'response.write(query)
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/02/01',102) and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/02/28',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input onkeyup="OnKeyUp_Verif(this)" type="text" value="<%=getvalueindice("ivalue")%>"name="value2" onkeypress ="javascript:CtrlVal('value2')"></td>
											<%
											else
											%>
												<td><input onkeyup="OnKeyUp_Verif(this)" type="text" name="value2" value=0  onkeypress ="javascript:CtrlVal('value2')"></td>
											<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/02/01',102) and CONVERT(DATETIME,'" & year(date) & "/02/28',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input onkeyup="OnKeyUp_Verif(this)" type="text" value="<%=getvalueindice("ivalue")%>"name="value2" onkeypress ="javascript:CtrlVal('value2')"></td><%
											else
										
												%>
												<td><input onkeyup="OnKeyUp_Verif(this)" type="text" name="value2" value=0 onkeypress ="javascript:CtrlVal('value2')"></td>
												<%
											end if
										end if	

										%>
								</tr>
								<tr>
									<td width="79">March</td>
									<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/03/01',102) and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/03/31',102)" 
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input onkeyup="OnKeyUp_Verif(this)" type="text" value="<%=getvalueindice("ivalue")%>"name="value3" onkeypress ="javascript:CtrlVal('value3')"></td>
												<%
											else
												%>
												<td><input onkeyup="OnKeyUp_Verif(this)" type="text" name="value3" value=0 onkeypress ="javascript:CtrlVal('value3')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/03/01',102) and CONVERT(DATETIME,'" & year(date) & "/03/31',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input onkeyup="OnKeyUp_Verif(this)"  type="text" value="<%=getvalueindice("ivalue")%>"name="value3" onkeypress ="javascript:CtrlVal('value3')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value3" value=0 onkeypress ="javascript:CtrlVal('value3')"></td>
												<%
											end if
										end if	

										%>

								</tr>
								<tr>
									<td width="79">April</td>
										<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/04/01',102) and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/04/30',102)" 
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value4" onkeypress ="javascript:CtrlVal('value4')"></td>
												<%
											else
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value4" value=0 onkeypress ="javascript:CtrlVal('value4')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/04/01',102) and CONVERT(DATETIME,'" & year(date) & "/04/30',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value4" onkeypress ="javascript:CtrlVal('value4')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value4" value=0 onkeypress ="javascript:CtrlVal('value4')"></td>
												<%
											end if
										end if	
										%>								
								</tr>
								<tr>
									<td width="79">May</td>
									<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/05/01',102)and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/05/31',102)" 
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value5" onkeypress ="javascript:CtrlVal('value5')"></td>
												<%
											else
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value5" value=0 onkeypress ="javascript:CtrlVal('value5')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "'  And idate between CONVERT(DATETIME,'" &year(date) & "/05/01',102) and CONVERT(DATETIME,'" & year(date) & "/05/31',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value5" onkeypress ="javascript:CtrlVal('value5')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value5" value=0 onkeypress ="javascript:CtrlVal('value5')"></td>
												<%
											end if
										end if	
										%>
								
								</tr>
								<tr>
									<td width="79">June</td>
									<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/06/01',102)and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/06/30',102)" 
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value6" onkeypress ="javascript:CtrlVal('value6')"></td>
												<%
											else
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value6" value=0 onkeypress ="javascript:CtrlVal('value6')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/06/01',102) and CONVERT(DATETIME,'" & year(date) & "/06/30',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value6" onkeypress ="javascript:CtrlVal('value6')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value6" value=0 onkeypress ="javascript:CtrlVal('value6')"></td>
												<%
											end if
										end if	
										%>
								
								</tr>
								<tr>
									<td width="79">July</td>
										<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/07/01',102)and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/07/31',102)" 
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value7" onkeypress ="javascript:CtrlVal('value7')"></td>
												<%
											else
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value7" value=0 onkeypress ="javascript:CtrlVal('value7')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/07/01',102) and CONVERT(DATETIME,'" & year(date) & "/07/31',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value7" onkeypress ="javascript:CtrlVal('value7')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value7" value=0 onkeypress ="javascript:CtrlVal('value7')"></td>
												<%
											end if
										end if	
	
										%>
								</tr>
								<tr>
									<td width="79">August</td>
										<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/08/01',102) and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/08/31',102)" 
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value8" onkeypress ="javascript:CtrlVal('value8')"></td>
												<%
											else
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value8" value=0 onkeypress ="javascript:CtrlVal('value8')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/08/01',102) and CONVERT(DATETIME,'" & year(date) & "/08/31',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value8" onkeypress ="javascript:CtrlVal('value8')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value8" value=0 onkeypress ="javascript:CtrlVal('value8')"></td>
												<%
											end if
										end if	

										%>
								</tr>
								<tr>
									<td width="79">September</td>
									<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/09/01',102)and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/09/30',102)" 
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value9" onkeypress ="javascript:CtrlVal('value9')"></td>
												<%
											else
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value9" value=0 onkeypress ="javascript:CtrlVal('value9')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/09/01',102) and CONVERT(DATETIME,'" & year(date) & "/09/30',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value9" onkeypress ="javascript:CtrlVal('value9')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value9" value=0 onkeypress ="javascript:CtrlVal('value9')"></td>
												<%
											end if
										end if	

										%>
 								</tr>
								<tr>
									<td width="79">October</td>
										<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/10/01',102)and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/10/31',102)" 
									set getvalueindice=connection.execute(query)
									if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value10" onkeypress ="javascript:CtrlVal('value10')"></td>
												<%
											else
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value10" value=0 onkeypress ="javascript:CtrlVal('value10')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/10/01',102) and CONVERT(DATETIME,'" & year(date) & "/10/31',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value10" onkeypress ="javascript:CtrlVal('value10')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value10" value=0 onkeypress ="javascript:CtrlVal('value10')"></td>
												<%
											end if
										end if	
										%>
								</tr>
								<tr>
									<td width="79" height="25">November</td>
										<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/11/01',102)and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/11/30',102)" 
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td height="25"><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value11" onkeypress ="javascript:CtrlVal('value11')"></td>
												<%
											else
												%>
												<td height="25"><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value11" value=0 onkeypress ="javascript:CtrlVal('value11')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/11/01',102) and CONVERT(DATETIME,'" & year(date) & "/11/30',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value11" onkeypress ="javascript:CtrlVal('value11')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value11" value=0 onkeypress ="javascript:CtrlVal('value11')"></td>
												<%
											end if
										end if	
	
										%>

								</tr>
								<tr>
									<td width="79">December</td>
										<%
										if not(request.querystring("annee")="") then
											query="select ivalue from Adv_Indices_values where IdIndice='" & request.querystring("id_indice0")& "' And idate  between CONVERT(DATETIME,'" &trim(request.querystring("annee0")) & "/12/01',102)and CONVERT(DATETIME,'" & trim(request.querystring("annee0")) & "/12/31',102)" 
											set getvalueindice=connection.execute(query)
											'response.write(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value12" onkeypress ="javascript:CtrlVal('value12')"></td>
												<%
											else
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value12" value=0 onkeypress ="javascript:CtrlVal('value12')"></td>
												<%
											end if
										else
											query="select ivalue from Adv_Indices_values where IdIndice='" & GetIndice("IdIndice") & "' And idate between CONVERT(DATETIME,'" &year(date) & "/12/01',102) and CONVERT(DATETIME,'" & year(date) & "/12/31',102)"
											set getvalueindice=connection.execute(query)
											if not getvalueindice.eof then
												%><td><input type="text" onkeyup="OnKeyUp_Verif(this)" value="<%=getvalueindice("ivalue")%>"name="value12" onkeypress ="javascript:CtrlVal('value12')"></td><%
											else
										
												%>
												<td><input type="text" onkeyup="OnKeyUp_Verif(this)" name="value12" value=0 onkeypress ="javascript:CtrlVal('value12')"></td>
												<%
											end if
										end if	
	
										%>
								</tr>
						</table>
					</td>
					<td width="0">
					</td>
				</tr>
			</table>
			
				</td>			
			</tr>
			<tr>
			<td>
				
				<input type="button" value="Save" style="float: right" onclick="verifier();">	
			</td>
			</tr>
		</table>		
		
				</p>
		
		</form>
		
		
		
		
	
		
	</body>
		</html>