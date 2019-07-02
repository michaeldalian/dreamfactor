<html>
	<head>
		<script language="javascript">
			<!--#INCLUDE FILE="connection/connection.asp"-->
			function CtrlVal(Nom)
			{
				// Procedure de validation
				if(window.event.keyCode==46 || window.event.keyCode==44)
				{
					window.event.keyCode = 44;
				}				
				if (window.event.keyCode<48 || window.event.keyCode>57)
				{
					if(window.event.keyCode=44)
					{
						window.event.keyCode=44;// On autorise la frappe de la touche 46 qui est inférieur à 48 qui est en fait le .
					}
					else
					window.event.keyCode = 0;
				}
			}
			function verifier()
			{
				var validation=true
				if(document.valid.value4.value=="" && document.valid.value5.value=="" && document.valid.value6.value=="" && document.valid.value7.value=="" && document.valid.value8.value=="" && document.valid.value9.value=="" && document.valid.value10.value=="" && document.valid.value11.value=="")
				{
					alert("All values can't be null")
					validation=false
				}
				//var maReg = new RegExp ( /[0-9]{0,1}[,]{0,1}[0-9]/ ) ;
				var maReg = new RegExp (/\d+\,?\d*|\,\d+/);

			 	for(i=4;i<12;i++)
			 	{
			 		var Elt=document.getElementById('value'+i)
			 		if (Elt.value!='' && Elt.value.search(maReg ) == -1  && Elt.value!=0)
						{
							alert ( "The 2003 value is not at the good format" ) ;
							validation=false
						}
			 	}
				if(validation ==true)
				{
					//supp_blanc();
					document.valid.submit();
				}
			}
			function changer()
			{
				
				document.debut.submit();
				document.valid.id_currency.value=document.debut.id_currency.value;
			}
			function chargement()
			{
				document.valid.id_currency.value=document.debut.id_currency.value;
				//supp_blanc();		
			}
			function supp_blanc()
			{
					for(i=4;i<=11;i++)
				{
					var Elt=document.getElementById('value'+i)
      				var newvalue=Elt.value
      				var exp=new RegExp(" ","g");
      				newvalue=newvalue.replace(exp, "");
      				Elt.value=newvalue; 
				}
			}
			OnKeyUp_Verif=function(obj) 
			{
   				var virgule=false;
   				var oldvalue=obj.value;
   				var newvalue=""
   				for (var i=0; i<oldvalue.length; i++) 
   				{
      				if ((oldvalue.charAt(i) == "," || oldvalue.charAt(i) == ".") && !virgule) 
      				{
         				if (i==0) 
         				{
            				newvalue+="0"
         				}
         				newvalue+=","
         				virgule=true;
      				}
      				else if (!isNaN(parseInt("" + oldvalue.charAt(i) + ""))) 
      				{
         				newvalue+=oldvalue.charAt(i)
      				}
   				}
   				if (oldvalue!=newvalue) 
   				{
      				obj.value=newvalue;
   				}
			}

		</script>
		<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
		<link rel="stylesheet" type="text/css" href="http://intra-dev/intrapur/procurement_mrs/css/log.css">
		<title>Currency Rate Update</title>
		<style>
			all.clsMenuItemNS{font: bold x-small Verdana; color: white; text-decoration: none;}
			.clsMenuItemIE{text-decoration: none; font: bold xx-small Verdana; color: white; cursor: hand;}
			A:hover {color: red;}
		</style>
	</head>
	<body onload="chargement();" bgcolor="#FFFFFF" text="#000000" background="http://intra-dev/intrapur/procurement_mrs/images/DefaultHomePage.gif" style="background-repeat: no-repeat;background-attachment: scroll;background-position: center center">
		<p align="center">
		<% 
 
		'Permet d'afficher tous les parametres reçus
			'For Each chaine_requete In Request.QueryString
  				'Response.Write "<tr>;<td>" & chaine_requete & "</td> <td>= " & Request.QueryString(chaine_requete) & "</td></tr>"       
			'Next
		if (request.querystring("update")=1 ) then
				if not(request.querystring("value4")="") then
					'value=Replace(request.querystring("Value4")," ","")
					'response.write("Nouvelle valeur" & value)
					ps=" exec ps_adv_MajCurrency_Rate" & " '" & request.querystring("id_currency") & "',2003," & Replace(request.querystring("value4"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value5")="") then
					ps=ps &"exec ps_adv_MajCurrency_Rate" & " '" & request.querystring("id_currency") & "',2004," & Replace(request.querystring("value5"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value6")="") then
					ps=ps &"exec ps_adv_MajCurrency_Rate" & " '" & request.querystring("id_currency") & "',2005," & Replace(request.querystring("value6"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value7")="") then
					ps=ps &"exec ps_adv_MajCurrency_Rate" & " '" & request.querystring("id_currency") & "',2006," & Replace(request.querystring("value7"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value8")="") then
					ps=ps &"exec ps_adv_MajCurrency_Rate" & " '" & request.querystring("id_currency") & "',2007," & Replace(request.querystring("value8"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value9")="") then
					ps=ps &"exec ps_adv_MajCurrency_Rate" & " '" & request.querystring("id_currency") & "',2008," & Replace(request.querystring("value9"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value10")="") then
					ps=ps &"exec ps_adv_MajCurrency_Rate" & " '" & request.querystring("id_currency") & "',2009," & Replace(request.querystring("value10"),",",".")  
					ps=ps & ",'';"
				end if
				if not(request.querystring("value11")="") then
					ps=ps &"exec ps_adv_MajCurrency_Rate" & " '" & request.querystring("id_currency") & "',2010," & Replace(request.querystring("value11"),",",".")  
					ps=ps & ",'';"
				end if
				'response.write(ps)
				set update_indice=connection.Execute(ps)
				response.write("Update has been done correctly")
			end if	
		%>
		<p align="center">
			<table width="472">
				<tr>
					<td align="center" class="titrelog">
						<p align="center">
						<span style="font-weight: 400">
						<font size="5">Currency Rate Update</font> </u>
						</font>
						</span>
						<p align="center">&nbsp;
					</td>			
				</tr>
				<tr>
					<td align="center">
						<table border="0" id="table1">
						<form name="debut" action="ASRAM_update_Currency.asp" method="get">
							<tr>
								<td valign="top"class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
									<b><font size="2">Currency</font><font size="4">
								</td> 	
								<td valign="middle" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">
									<select name="id_currency" onchange="changer();">
									<%	
									set GetCurrency=connection.Execute("select * from Adv_Currency")
									if not GetCurrency.eof then
										do until GetCurrency.eof
											if(trim(request.querystring("id_currency"))=trim(GetCurrency("IdCurrency"))) then
											%>
												<OPTION selected value="<%=trim(GetCurrency("IdCurrency"))%>"> <%=GetCurrency("IdCurrency")%> </OPTION>
											<%else%>
												<OPTION value="<%=trim(GetCurrency("IdCurrency"))%>"> <%=GetCurrency("IdCurrency")%></OPTION>
											<%
											end if
											GetCurrency.movenext
										loop
									end if
									%>
								</select>
								</form>								</td>
							</tr>
					<tr>
					<form name="valid" action="ASRAM_update_Currency.asp" method="get">
						<input type=hidden name="id_currency">
						<input type=hidden name="update" value=1>
						<td valign="top"class="titrelog" style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#DCE9F6">
							<span style="font-weight: 400">
							<font size="2"><b>Rate</b></font></span></td>
						<td style="border: 1px solid #000000; padding-left: 4px; padding-right: 4px; padding-top: 1px; padding-bottom: 1px" bgcolor="#FFFFFF">
							<table border="0" width="100%" id="table2">
								<tr>
									<td width="79">2003</td>
										<%
										'response.write(query)
										if not(request.querystring("id_currency")="") then
											query="select rate from Adv_Currency_Rate where IdCurrency='" & request.querystring("id_currency")& "' and iyear=2003"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" value="<%=get_rate("Rate")%>"name="value4" onkeyup="OnKeyUp_Verif(this)" onkeypress ="javascript:CtrlVal('value4')"></td>
												<%
											else
												%>
												<td><input type="float" name="value4" value=0 onkeyup="OnKeyUp_Verif(this)" onkeypress ="javascript:CtrlVal('value4')"></td>
												<%
											end if
										else
											set Getcurrency=connection.Execute("select top 1 IdCurrency from Adv_Currency order by IdCurrency")
											query="select rate from Adv_Currency_Rate where IdCurrency='" &Getcurrency("IdCurrency")& "' and iyear=2003"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value="<%=get_rate("Rate")%>"name="value4" onkeypress ="javascript:CtrlVal('value4')"></td><%
											else
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value=0 name="value4" onkeypress ="javascript:CtrlVal('value4')"></td><%
											end if
										end if
	
										%>
								</tr>
																<tr>
									<td width="79">2004</td>
										<%
										'response.write(query)
										if not(request.querystring("id_currency")="") then
											query="select rate from Adv_Currency_Rate where IdCurrency='" & request.querystring("id_currency")& "' and iyear=2004"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value="<%=get_rate("Rate")%>"name="value5" onkeypress ="javascript:CtrlVal('value5')"></td>
												<%
											else
												%>
												<td><input type="float" onkeyup="OnKeyUp_Verif(this)"name="value5" value=0  onkeypress ="javascript:CtrlVal('value5')"></td>
												<%
											end if
										else
											set Getcurrency=connection.Execute("select top 1 IdCurrency from Adv_Currency order by IdCurrency")
											query="select rate from Adv_Currency_Rate where IdCurrency='" &Getcurrency("IdCurrency")& "' and iyear=2004"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float"onkeyup="OnKeyUp_Verif(this)" value="<%=get_rate("Rate")%>"name="value5" onkeypress ="javascript:CtrlVal('value5')"></td><%
											else
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value=0 name="value5" onkeypress ="javascript:CtrlVal('value5')"></td><%
											end if
										end if
	
										%>
								</tr>
																<tr>
									<td width="79">2005</td>
										<%
										'response.write(query)
										if not(request.querystring("id_currency")="") then
											query="select rate from Adv_Currency_Rate where IdCurrency='" & request.querystring("id_currency")& "' and iyear=2005"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value6" onkeypress ="javascript:CtrlVal('value6')"></td>
												<%
											else
												%>
												<td><input type="float" onkeyup="OnKeyUp_Verif(this)"name="value6" value=0  onkeypress ="javascript:CtrlVal('value6')"></td>
												<%
											end if
										else
											set Getcurrency=connection.Execute("select top 1 IdCurrency from Adv_Currency order by IdCurrency")
											query="select rate from Adv_Currency_Rate where IdCurrency='" &Getcurrency("IdCurrency")& "' and iyear=2005"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input onkeyup="OnKeyUp_Verif(this)"type="float" value="<%=get_rate("Rate")%>"name="value6" onkeypress ="javascript:CtrlVal('value6')"></td><%
											else
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value=0 name="value6" onkeypress ="javascript:CtrlVal('value6')"></td><%
											end if

										end if
	
										%>
								</tr>
																<tr>
									<td width="79">2006</td>
										<%
										'response.write(query)
										if not(request.querystring("id_currency")="") then
											query="select rate from Adv_Currency_Rate where IdCurrency='" & request.querystring("id_currency")& "' and iyear=2006"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value7" onkeypress ="javascript:CtrlVal('value7')"></td>
												<%
											else
												%>
												<td><input type="float" onkeyup="OnKeyUp_Verif(this)"name="value7" value=0  onkeypress ="javascript:CtrlVal('value7')"></td>
												<%
											end if
										else
											set Getcurrency=connection.Execute("select top 1 IdCurrency from Adv_Currency order by IdCurrency")
											query="select rate from Adv_Currency_Rate where IdCurrency='" &Getcurrency("IdCurrency")& "' and iyear=2006"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value7" onkeypress ="javascript:CtrlVal('value7')"></td><%
											else
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value=0 name="value7" onkeypress ="javascript:CtrlVal('value7')"></td><%
											end if

										end if
	
										%>
								</tr>
																<tr>
									<td width="79">2007</td>
										<%
										'response.write(query)
										if not(request.querystring("id_currency")="") then
											query="select rate from Adv_Currency_Rate where IdCurrency='" & request.querystring("id_currency")& "' and iyear=2007"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value8" onkeypress ="javascript:CtrlVal('value8')"></td>
												<%
											else
												%>
												<td><input type="float" onkeyup="OnKeyUp_Verif(this)"name="value8" value=0  onkeypress ="javascript:CtrlVal('value8')"></td>
												<%
											end if
										else
											set Getcurrency=connection.Execute("select top 1 IdCurrency from Adv_Currency order by IdCurrency")
											query="select rate from Adv_Currency_Rate where IdCurrency='" &Getcurrency("IdCurrency")& "' and iyear=2007"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value8" onkeypress ="javascript:CtrlVal('value8')"></td><%
											else
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value=0 name="value8" onkeypress ="javascript:CtrlVal('value8')"></td><%
											end if

										end if
										%>
								</tr>
																<tr>
									<td width="79">2008</td>
										<%
										'response.write(query)
										if not(request.querystring("id_currency")="") then
											query="select rate from Adv_Currency_Rate where IdCurrency='" & request.querystring("id_currency")& "' and iyear=2008"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value9" onkeypress ="javascript:CtrlVal('value9')"></td>
												<%
											else
												%>
												<td><input type="float" onkeyup="OnKeyUp_Verif(this)"name="value9" value=0  onkeypress ="javascript:CtrlVal('value9')"></td>
												<%
											end if
										else
											set Getcurrency=connection.Execute("select top 1 IdCurrency from Adv_Currency order by IdCurrency")
											query="select rate from Adv_Currency_Rate where IdCurrency='" &Getcurrency("IdCurrency")& "' and iyear=2008"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value9" onkeypress ="javascript:CtrlVal('value9')"></td><%
											else
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value=0 name="value9" onkeypress ="javascript:CtrlVal('value9')"></td><%
											end if

										end if

										%>
								</tr>
																<tr>
									<td width="79">2009</td>
										<%
										'response.write(query)
										if not(request.querystring("id_currency")="") then
											query="select rate from Adv_Currency_Rate where IdCurrency='" & request.querystring("id_currency")& "' and iyear=2009"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value10" onkeypress ="javascript:CtrlVal('value10')"></td>
												<%
											else
												%>
												<td><input type="float" onkeyup="OnKeyUp_Verif(this)"name="value10" value=0  onkeypress ="javascript:CtrlVal('value10')"></td>
												<%
											end if
										else
											set Getcurrency=connection.Execute("select top 1 IdCurrency from Adv_Currency order by IdCurrency")
											query="select rate from Adv_Currency_Rate where IdCurrency='" &Getcurrency("IdCurrency")& "' and iyear=2009"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value10" onkeypress ="javascript:CtrlVal('value10')"></td><%
											else
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value=0 name="value10" onkeypress ="javascript:CtrlVal('value10')"></td><%
											end if
										end if

										%>
								</tr>
																<tr>
									<td width="79">2010</td>
										<%
										'response.write(query)
										if not(request.querystring("id_currency")="") then
											query="select rate from Adv_Currency_Rate where IdCurrency='" & request.querystring("id_currency")& "' and iyear=2010"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value11" onkeypress ="javascript:CtrlVal('value11')"></td>
												<%
											else
												%>
												<td><input type="float" onkeyup="OnKeyUp_Verif(this)"name="value11" value=0  onkeypress ="javascript:CtrlVal('value11')"></td>
												<%
											end if
										else
											set Getcurrency=connection.Execute("select top 1 IdCurrency from Adv_Currency order by IdCurrency")
											query="select rate from Adv_Currency_Rate where IdCurrency='" &Getcurrency("IdCurrency")& "' and iyear=2010"
											set get_rate=connection.execute(query)
											if not get_rate.eof then
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)"value="<%=get_rate("Rate")%>"name="value11" onkeypress ="javascript:CtrlVal('value11')"></td><%
											else
												%><td><input type="float" onkeyup="OnKeyUp_Verif(this)" value=0 name="value11" onkeypress ="javascript:CtrlVal('value11')"></td><%
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