
<html>
<head>


<script language="JavaScript">


function scroll(jumpSpaces,position) {
	var msg = " Welcome on Asram :"
	var out = ""

	if (killScroll) return false; 

	for (var i=0; i<position; i++){
		out += msg.charAt(i)
		}
	for (i=1;i<jumpSpaces;i++) {
		out += " "
		}

	out += msg.charAt(position)
	window.status = out

	if (jumpSpaces <= 1) {
		position++
		if (msg.charAt(position) == ' ') {
			position++ 
			}
		jumpSpaces = 100-position        
		}
	else if (jumpSpaces >  3) {
		jumpSpaces *= .75
		}
	else {
		jumpSpaces--
		}

	if (position != msg.length) {
		var cmd = "scroll(" + jumpSpaces + "," + position + ")";
		scrollID = window.setTimeout(cmd,5);
		} 
	else {
		scrolling = false
		return false
		}
	return true;
	}

function startScroller() {
	if (scrolling)
		if (!confirm('Re-initialize snapIn?')) return false

	killScroll = true
	scrolling = true
	var killID = window.setTimeout('killScroll=false',6)
	scrollID = window.setTimeout('scroll(100,0)',10)
	return true
	}

var scrollID = Object
var scrolling = false
var killScroll = false


window.onload = startScroller;

</SCRIPT>

<title>Asram</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<meta http-equiv="Page-Enter" content="blendTrans(Duration=2.0)">
<script language="javascript" src="../images/function.js"></script>
<style>
<!--
div.Section1
	{page:Section1;}
 p.MsoNormal
	{margin:0cm;
	margin-bottom:.0001pt;
	font-size:12.0pt;
	font-family:"Times New Roman";}
-->
</style>
</head>

<link rel="stylesheet" href="../css/log.css" type="text/css">
<body bgcolor="#FFFFFF" text="#000000">
<table width="75%" border="0">
  <tr> 
    <td height="130" width="23%">
    <IMG SRC="images/argent.gif" onMouseOver="ejs_img_fx(this)" onload="window.setTimeout('document.images[0].src=\'images/balance.bmp\'', 5000)">
    
</td>		
    <td height="22" width="23%">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" width="54%">&nbsp;</td>
  </tr>
  <tr> 
    <td height="22" width="23%"><i><a class="titrelog">Menu</a></i></td>
    <td height="22" width="54%">&nbsp;</td>
  </tr>
 </table>

<table>
  <tr> 
    
    <td>
    <a href="BusEdit_ASRaM.asp" target="mainFrame" class="rublog">
	<b>C</b>reate New File</a>
   </td>

  </tr>
</table>
<table>
  <tr> 
   

    <td>
	<a href="../Asram/FindAsram.asp" target="mainFrame" class="rublog">
	<b>M</b>odify Existing File</a></td>

  </tr>
</table>
<table>
  <tr> 
    <td>&nbsp;</td>

    <td>
</td>

  </tr>
</table>

<table>
 <tr><td>&nbsp;</td> </tr>
  <tr> 
        <td>
	<a href="../Asram/Managedata_Asram.asp" target="mainFrame" class="rublog">
	<b>M</b>anage Data</a></td>
    <td></td>

  </tr>
  <tr><td></td><td>&nbsp;</td></tr>
  
</table>
<!-- DEBUT DU SCRIPT -->
<STYLE TYPE="text/css">
.ejs_scroll {font-size:11px;font-family:Arial;color:#CC0000;}
</STYLE>
<table style="font-family: Times New Roman; font-size: 12pt; color: #000080">
  <tr><td>_-_-_-_-_-_-_-_-_</td><td></td></tr>
 <tr><td>ASRaM: Advance Savings Realisation and Monitoring</td> </tr>
  <tr><td>_-_-_-_-_-_-_-_-_</td><td>&nbsp;</td></tr>
  <tr><td></td><td>&nbsp;</td></tr>
  <tr><td></td><td>&nbsp;</td></tr>
</table>
<script language="JavaScript1.2">

ejs_scroll_largeur = 100;
ejs_scroll_hauteur = 125;
ejs_scroll_bgcolor = '#FFFFFF';
/* Mettre ici le chemin de l'image de fond */
ejs_scroll_background = "";
/* Mettre ici le temps en secondes */
ejs_scroll_pause_seconde = 5;

ejs_scroll_message = new Array;

ejs_scroll_message[0]='<a href="mailto:fabrice.demichelis@eurocopter.com"><SPAN class="rublog">Help and Support : Demichelis F. 59496</SPAN></FONT></a>';
ejs_scroll_message[1]='<a href="mailto:fabrice.demichelis@eurocopter.com"><SPAN class="rublog">Creation: Demichelis F. / Dalian M. </SPAN></FONT></a>';
function d(texte)
	{
	document.write(texte);
	}

d('<DIV ID=ejs_scroll_relativ STYLE="position:relative;width:'+ejs_scroll_largeur+';height:'+ejs_scroll_hauteur+';background-color:'+ejs_scroll_bgcolor+';background-image:url('+ejs_scroll_background+')">');
d('<DIV ID=ejs_scroll_cadre STYLE="position:absolute;width:'+(ejs_scroll_largeur-8)+';height:'+(ejs_scroll_hauteur-8)+';top:4;left:4;clip:rect(0 '+(ejs_scroll_largeur-8)+' '+(ejs_scroll_hauteur-8)+' 0)">');
d('<div id=ejs_scroller_1 style="position:absolute;width:'+(ejs_scroll_largeur-8)+';left:0;top:0;" CLASS=ejs_scroll>'+ejs_scroll_message[0]+'</DIV>');
d('<div id=ejs_scroller_2 style="position:absolute;width:'+(ejs_scroll_largeur-8)+';left:0;top:'+ejs_scroll_hauteur+';" CLASS=ejs_scroll>'+ejs_scroll_message[1]+'</DIV>');
d('</DIV></DIV>');

ejs_scroll_mode =1;
ejs_scroll_actuel = 0;

function ejs_scroll_start()
	{
	if(ejs_scroll_mode == 1)
		{
		ejs_scroller_haut = "ejs_scroller_1";
		ejs_scroller_bas = "ejs_scroller_2";
		ejs_scroll_mode = 0;
		}
	else
		{
		ejs_scroller_bas = "ejs_scroller_1";
		ejs_scroller_haut = "ejs_scroller_2";
		ejs_scroll_mode = 1;
		}
	ejs_scroll_nb_message = ejs_scroll_message.length-1;
	if(ejs_scroll_actuel == ejs_scroll_nb_message)
		ejs_scroll_suivant = 0;
	else
		ejs_scroll_suivant = ejs_scroll_actuel+1;
	if(document.getElementById)
		document.getElementById(ejs_scroller_bas).innerHTML = ejs_scroll_message[ejs_scroll_suivant];
	ejs_scroll_top = 0;
	if(document.getElementById)
		setTimeout("ejs_scroll_action()",ejs_scroll_pause_seconde*1000)
	}

function ejs_scroll_action()
	{
	ejs_scroll_top -= 1;
	document.getElementById(ejs_scroller_haut).style.top = ejs_scroll_top;
	document.getElementById(ejs_scroller_bas).style.top = ejs_scroll_top+ejs_scroll_hauteur;
	if((ejs_scroll_top+ejs_scroll_hauteur) > 0)
		setTimeout("ejs_scroll_action()",10)
	else
		ejs_scroll_stop()
	}

function ejs_scroll_stop()
	{
	ejs_scroll_actuel = ejs_scroll_suivant;
	ejs_scroll_start()
	}

window.onload = ejs_scroll_start;
</SCRIPT>



</body>
</html>