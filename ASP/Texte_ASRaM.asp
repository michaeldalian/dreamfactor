<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<title>GAP Analysis</title>
</head>

<body onunload="document.getElementById('TSave').click()">
<form action="Texte_ASRaM.asp" name="Texte_ASRaM.asp" method="post">
<%
'affiche toute les valeurs envoyée avec la methode post
response.write "<table>"
      For Each chaine_requete In Request.Form
       Response.Write "<tr><td>///" & chaine_requete & "///</td><td> = " _
                   & Request.Form(chaine_requete) & "/// </td></tr>"
     Next
  response.write "</table>"      

IdBusCase = request("IdBusCase")
Comments = request("Comments")
NomFichier=Server.MapPath("Textes/"&IdBusCase&".txt")
Set FSO = Server.CreateObject("Scripting.FileSystemObject")

if Request("TSave") = "Save" then
	set Fichier=nothing

Set Fichier = FSO.CreateTextFile(NomFichier, True)
	'---écriture dans ce fichier
	Fichier.write(Comments)
	Fichier.close
	%>
	<script language="javascript">
	window.opener.location.reload()
	window.close()
	</script>
<%	
end if
'----------teste si le fichier des commentaires existe------------

If FSO.FileExists(NomFichier) Then
	response.write "le fichier existe bien sur le serveur"
	Set Fichier = FSO.OpenTextFile(NomFichier)
Else
	'-------------Création du fichier---------------
	Set Fichier = FSO.CreateTextFile(NomFichier)
end if
	if not (Fichier.atEndOfStream) then Comments = Fichier.ReadAll
	
	set Fichier=nothing
	Set FSO = Nothing
%>

<input type="hidden" name="IdBusCase" size="13" value ="<%=IdBusCase%>">
<table border="2" width="100%" style="border-style:ridge; border-width:3px; " bordercolorlight="#00D2A8" bordercolordark="#00A882">
<tr><td height="5" align="center" width="54%">GAP Analysis</td></tr>
<tr><td colspan="2" align="center"><textarea cols="118" name="Comments" rows="7" style="border-style: inset; border-width: 2px;">
<%=Comments%></textarea>
</td></tr>
<tr><td height="5"><input type="submit" value="Save" name="TSave" style="padding: 0; color:#008080">
</td></tr>

</table>

</form>
</body>

</html>
