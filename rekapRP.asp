<%
'***** START WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ****** ' 
'Application: HEALTH WEALTH INTERNATIONAL
'Author: Ir. Peter Sindoro
'Coding : August 2009
'Revised : - 
'***** END WARNING - REMOVAL OR MODIFICATION OF THIS CODE WILL VIOLATE THE LICENSE AGREEMENT ****** 
%>
<!--#Include File=pg_config.asp-->
<%
judulmenu = session("menuasal")
session("menuasal") = judulmenu
session("menu_bag") = session("menu_bag")
session("tema") = "home"
babe = session("mimin")
sort = request("sort")
txtcari = request("keyword")
itung = request("sesi")

if session("mimin") = "" then
	session("error") = "Session login anda sudah expired, silahkan login kembali"
	response.redirect "default.asp"
else
	session("mimin") = babe	
end if
session("menu_id") = request("menu_id")	

%>
<!--#Include File=dbcon/opendbALL.asp-->
<!--#Include File=pg_cek_user.asp-->
<!--#Include File=dbcon/opendb.asp-->
<%
dim sapa
rs.Open "SELECT * FROM owner WHERE kta = '"&babe&"'",conn
	if rs.bof then
		rs.close
		set rs=nothing
		Conn.Close
		set conn=nothing
		session("error1") = "Admin id tidak ditemukan"
		response.redirect "default.asp"
	else		
		sapa = rs("nama")
		
	end if						
rs.close

tglskr = date()
thn = year(tglskr)
bln = month(tglskr)
bulan = request("bulan")
tahun = request("tahun")

wulan = bulan
nahun = tahun

if bln = 1 then
	blnskr = 12
	thnskr = thn-1	
else
	blnskr = bln-1
	thnskr = thn
end if	


if bulan <> "" and tahun <>"" then

rs.Open "SELECT count(id) FROM rep_bns_mycourse where bulan = '"&bulan&"' and tahun='"&tahun&"'",conn	
	if rs.bof then
		x = 0
	else
		x = cint(rs("count(id)"))
	end if
rs.Close

pgview = request("pgview")
if pgview = "" then
	bg = 1500
else
	bg = pgview
end if

pg = 0
pgas = request("page")
sort = request("sort")
if pgas = "" then
	pg = 0
else
if pgas<>"" then
	pg = pgas-1
end if		
end if

function roundup(x)
If x > Int(x) then
roundup = Int(x) + 1
Else
roundup = x
End If
End Function

pgcunt = x / bg

if pgcunt < 1 then
	pgct = 1
else
	pgct = roundup(pgcunt)
end if	


end if
%>

<html>

<head>
<meta http-equiv="Content-Language" content="en-us">
<meta http-equiv="Content-Type" content="text/html;charset=iso-8859-1">
<TITLE>Health Wealth International :: Rekap Perhitungan Bonus Bulanan</TITLE>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=windows-1252">
<style type="text/css">
<!--
td {
	font-family:Tahoma;
	font-size:11px;
	color:#4F4F4F;
}
a {
	text-decoration: none;
	color:#31509F;
}
a.1 {
	text-decoration: none;
	color:#BD1818;
}
a.2 {
	text-decoration: underline;
	color:#ffffff;
}
a.3 {
	text-decoration: underline;
	color:#FFD16F;
}
a.4 {
	text-decoration: none;
	color:#FFFFFF;
}
a.5 {
	text-decoration: underline;
	color:#BD1818;
}
.t11 {
	font-family: Tahoma;
	font-size: 11px;
	font-style: normal;
}
.t10 {
	font-family: Tahoma;
	font-size: 10px;
	font-style: normal;
}
.style1 {color: #FFFFFF}
.style2 {color: #616161}
.style3 {color: #31509F}
.style4 {color: #ffffff}

-->
</style>
<style fprolloverstyle>A:hover {font-size: 11px; color: #FF9966; font-family: Tahoma; text-decoration: underline}
</style>

<style type="text/css">
.pagination{
padding: 2px;
}

.pagination ul{
margin: 0;
padding: 0;
text-align: left; /*Set to "right" to right align pagination interface*/
font-size: 11px;
}

.pagination li{
list-style-type: none;
display: inline;
padding-bottom: 1px;
}

.pagination a, .pagination a:visited{
padding: 0 5px;
border: 1px solid #9aafe5;
text-decoration: none; 
color: #2e6ab1;
}

.pagination a:hover, .pagination a:active{
border: 1px solid #2b66a5;
color: #000;
background-color: #FFFF80;
}

.pagination a.currentpage{
background-color: #2e6ab1;
color: #FFF !important;
border-color: #2b66a5;
font-weight: bold;
cursor: default;
}

.pagination a.disablelink, .pagination a.disablelink:hover{
background-color: white;
cursor: default;
color: #929292;
border-color: #929292;
font-weight: normal !important;
}

.pagination a.prevnext{
font-weight: bold;
}
</style>
<SCRIPT language="JavaScript1.2" src="gen_validation.js"></SCRIPT>

<script language="javascript">
var win = null;
function NewWindow(mypage,myname,w,h,scroll){
LeftPosition = (screen.width) ? (screen.width-w)/2 : 0;
TopPosition = (screen.height) ? (screen.height-h)/2 : 0;
settings =
'height='+h+',width='+w+',top='+TopPosition+',left='+LeftPosition+',scrollbars='+scroll+',resizable'
win = window.open(mypage,myname,settings)
}
</script>

</head>

<body topmargin="0" leftmargin="0" background="images/bg.gif">

<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td><!--#Include File=pg_header.asp--></td>
	</tr>
</table>

<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td width="18%" valign="top">
		<table border="0" cellpadding="0" cellspacing="0" width="100%">
			<tr>
				<td valign="top"><!--#Include File=pg_leftmenu.asp--></td>
			</tr>
			<tr>
				<td valign="top">&nbsp;</td>
			</tr>
			<tr>
				<td valign="top">&nbsp;</td>
			</tr>
			<tr>
				<td valign="top">&nbsp;</td>
			</tr>
		</table>
		</td>
		<td valign="top" width="82%">
		<div align="center">
			<table border="0" cellpadding="0" cellspacing="0" width="98%">
				<tr>
					<td valign="top" height="3"></td>
				</tr>
				<tr>
					<td valign="top">
					<table border="1" cellspacing="1" width="100%" style="border-collapse: collapse" bordercolor="#DBDBDB">
						<tr>
							<td valign="top">
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td bgcolor="#808080" height="20">
									<p align="center">
									<font size="3" color="#FFFFFF"><b>REDEMPTION 
									POINT</b></font></td>
								</tr>
							</table>
							<table border="0" cellpadding="0" cellspacing="0" width="100%">
								<tr>
									<td>&nbsp;</td>
								</tr>
								<tr>
									<td>
									<div align="center">
										<table border="0" cellpadding="0" cellspacing="0" width="98%">
											<tr>
												<td valign="top">
												<form name="pointtop" method="post" action="rekaprp.asp?menu_id=1" onsubmit="return formCheck1(this)">
												<b><font color="#000000">Input 
												Id Distributor : </font></b> <input type="text" name="ktana" style="font-size: 8pt; font-family: Verdana; border: 1px solid #808080" size="10">
												<input type="submit" name="btklaim" value="Lihat...!" style="font-size: 8pt; font-family: Verdana"></form><br>
											<% '''''''
											ktana = SafeSQL(ucase(request("ktana")))
											idambil = request("idambil")
											ket = ucase(request("ket"))
											idklaim = request("idklaim")

											if idambil <> "" then
												rs.Open "SELECT * FROM healthwealthint_com_hwi.tab_rp_sponsor_release where id like '"&idambil&"' ",conn,3,3
												if rs.bof then
												else
												rs.update
												rs("ambil") = "T"
												rs.update
												end if
												rs.close
											end if
											if idklaim <> "" then
												rs.Open "SELECT * FROM healthwealthint_com_hwi.tab_rp_sponsor_release where id like '"&idklaim&"' ",conn,3,3
												if rs.bof then
												else
												rs.update
												rs("ket") = ket
												rs.update
												end if
												rs.close
											%><b>&nbsp;&nbsp;
												<font color="#FF0000">UPDATE 
												KETERANGAN PENGAMBILAN BERHASIL 
												: <%=ket%></font></b><%	
											end if
											if ktana <> "" then
											%>											
												<table border="0" cellpadding="0" cellspacing="0" width="99%" height="284">
																<tr>
																	<td valign="top">
														<font size="2" color="#000000"><b>
														Redemption Point :</b></font></td>
																</tr>
																<tr>
																	<td valign="top">
														<%			 rs.Open "SELECT uid FROM healthwealthint_com_hwi.member where kta like '"&ktana&"' ",conn
												if rs.bof then
												uid = "-"
												else
												uid = ucase(rs("uid"))
												end if
												rs.close %>

																	&nbsp;<font color="#000000" size="2"><b><%=ktana%> 
														[ <%=uid%> ]</b></font></td>
																</tr>
																<tr>
																	<td valign="top">
																	<table border=2 cellpadding=10><tr><td>	
																	<b>
																	<font face="Calibri" size="4" color="#008000">
																	PERIODE 
																	JAN-SEP 2018</font></b><font color="#000000"><br>
												<b>Point yang telah anda 
																	dapatkan 
																	adalah 
																	sebanyak </b>
												</font><b>
												<% rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2018-01' and tgl < '2018-09-32' order by id ",conn
												if rs.bof then
												point = 0
												else
												point = rs("sum(point)")
												end if
												rs.close
												
												rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2018-01-12' and tgl < '2018-10-09' order by id ",conn
												if rs.bof then
												pointo = 0
												else
												pointo = rs("sum(point)")
												end if
												rs.close
												
												if isnull(pointo) = true then
												pointo = 0
												end if
												ponta = point - pointo
												%>
												<font color="#FF0000"><%=point%> </font>
																	,</b><font color="#000000"><b> 
																	dan telah 
																	klaim <font color="#FF0000"><%=pointo%> </font>
																	, jadi sisa <font color="#FF0000"><%=ponta%> </font> 
																	Point RP</b><br>
												<b>History Sponsorisasi anda :
												<% rs.Open "SELECT direk,paket,point,tgl FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2018-01' and tgl < '2018-09-32' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Paket</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("paket"))%> - <%=ucase(rs("direk"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>
<b>History Klaim anda :
												<% rs.Open "SELECT * FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2018-01-17' and tgl < '2018-10-17' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Keterangan</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Sudah Cetak</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
idklaim = rs("id")
ket = rs("ket")
ambil = rs("ambil")

if ambil = "F" then
warna = "white"
else
warna = "yellow"
end if

%>			
<tr bgcolor="<%=warna%>"><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("hadiah"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td><td align="center"><form name="lrekap4<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idklaim" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<font color="#000000"><b>pengambilan : </b></font>
				<input name="ket" value="<%=ket%>"style="font-size: 8pt; font-family: Verdana; border: 1px solid #808080; font-weight:700" size="40"><font color="#000000"><b>
				</b></font> 
				<input type="submit" name="btpoint" value="Update" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form></td><td align="center"><font color="#000000"><b><%if warna = "white" then%></b></font><form name="updrdc4<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idambil" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<input type="submit" name="btpoint" value="Sudah" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form><font color="#000000"><b><%end if%></b></font></td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table><b>Yang sudah di Ambil : 
<% rs.Open "SELECT * FROM healthwealthint_com_hwi.fx_ge_det where nopo like '%"&ktana&"%' ",conn,3,3 %>
												</b>
																	</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>No invoice</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Jumlah</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("nopo"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("kode"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("jumlah"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>

<br>&nbsp;</font><p></td></tr></table>
																	<table border=2 cellpadding=10><tr><td>	
																	<b>
																	<font face="Calibri" size="4" color="#008000">
																	PERIODE 
																	JUL-DES 2017</font></b><font color="#000000"><br>
												<b>Point yang telah anda 
																	dapatkan 
																	adalah 
																	sebanyak </b>
												</font><b>
												<% rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2017-07' and tgl < '2017-12-32' order by id ",conn
												if rs.bof then
												point = 0
												else
												point = rs("sum(point)")
												end if
												rs.close
												
												rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2017-07-12' and tgl < '2018-01-17' order by id ",conn
												if rs.bof then
												pointo = 0
												else
												pointo = rs("sum(point)")
												end if
												rs.close
												
												if isnull(pointo) = true then
												pointo = 0
												end if
												ponta = point - pointo
												%>
												<font color="#FF0000"><%=point%> </font>
																	,</b><font color="#000000"><b> 
																	dan telah 
																	klaim <font color="#FF0000"><%=pointo%> </font>
																	, jadi sisa <font color="#FF0000"><%=ponta%> </font> 
																	Point RP</b><br>
												<b>History Sponsorisasi anda :
												<% rs.Open "SELECT direk,paket,point,tgl FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2017-07' and tgl < '2017-12-32' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Paket</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("paket"))%> - <%=ucase(rs("direk"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>
<b>History Klaim anda :
												<% rs.Open "SELECT * FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2017-07-12' and tgl < '2018-01-17' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Keterangan</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Sudah Cetak</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
idklaim = rs("id")
ket = rs("ket")
ambil = rs("ambil")

if ambil = "F" then
warna = "white"
else
warna = "yellow"
end if

%>			
<tr bgcolor="<%=warna%>"><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("hadiah"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td><td align="center"><form name="lrekap3<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idklaim" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<font color="#000000"><b>pengambilan : </b></font>
				<input name="ket" value="<%=ket%>"style="font-size: 8pt; font-family: Verdana; border: 1px solid #808080; font-weight:700" size="40"><font color="#000000"><b>
				</b></font> 
				<input type="submit" name="btpoint" value="Update" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form></td><td align="center"><font color="#000000"><b><%if warna = "white" then%></b></font><form name="updrdc3<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idambil" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<input type="submit" name="btpoint" value="Sudah" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form><font color="#000000"><b><%end if%></b></font></td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table><b>Yang sudah di Ambil : 
<% rs.Open "SELECT * FROM healthwealthint_com_hwi.fx_ge_det where nopo like '%"&ktana&"%' ",conn,3,3 %>
												</b>
																	</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>No invoice</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Jumlah</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("nopo"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("kode"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("jumlah"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>

<br>&nbsp;</font><p></td></tr></table>
																	<table border=2 cellpadding=10><tr><td>	
																	<b>
																	<font face="Calibri" size="4" color="#008000">
																	PERIODE 
																	JAN-JUN 2017</font></b><font color="#000000"><br>
												<b>Point yang telah anda 
																	dapatkan 
																	adalah 
																	sebanyak </b>
												</font><b>
												<% rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2017-01' and tgl < '2017-06-32' order by id ",conn
												if rs.bof then
												point = 0
												else
												point = rs("sum(point)")
												end if
												rs.close
												
												rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2017-01-06' and tgl < '2017-07-12' order by id ",conn
												if rs.bof then
												pointo = 0
												else
												pointo = rs("sum(point)")
												end if
												rs.close
												
												if isnull(pointo) = true then
												pointo = 0
												end if
												ponta = point - pointo
												%>
												<font color="#FF0000"><%=point%> </font>
																	,</b><font color="#000000"><b> 
																	dan telah 
																	klaim <font color="#FF0000"><%=pointo%> </font>
																	, jadi sisa <font color="#FF0000"><%=ponta%> </font> 
																	Point RP</b><br>
												<b>History Sponsorisasi anda :
												<% rs.Open "SELECT direk,paket,point,tgl FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2017-01' and tgl < '2017-06-32' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Paket</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("paket"))%> - <%=ucase(rs("direk"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>
<b>History Klaim anda :
												<% rs.Open "SELECT * FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2017-01-06' and tgl < '2017-07-12' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Keterangan</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Sudah Cetak</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
idklaim = rs("id")
ket = rs("ket")
ambil = rs("ambil")

if ambil = "F" then
warna = "white"
else
warna = "yellow"
end if

%>			
<tr bgcolor="<%=warna%>"><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("hadiah"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td><td align="center"><form name="lrekap<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idklaim" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<font color="#000000"><b>pengambilan :  </b></font>
				<input name="ket" value="<%=ket%>"style="font-size: 8pt; font-family: Verdana; border: 1px solid #808080; font-weight:700" size="40"><font color="#000000"><b>
				</b></font> 
				<input type="submit" name="btpoint" value="Update" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form></td><td align="center"><font color="#000000"><b><%if warna = "white" then%></b></font><form name="updrdc2<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idambil" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<input type="submit" name="btpoint" value="Sudah" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form><font color="#000000"><b><%end if%></b></font></td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table><b>Yang sudah di Ambil :
<% rs.Open "SELECT * FROM healthwealthint_com_hwi.fx_ge_det where nopo like '%"&ktana&"%' ",conn,3,3 %>
												</b>
																	</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>No invoice</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Jumlah</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("nopo"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("kode"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("jumlah"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>

<br>&nbsp;</font><p></td></tr></table>
																	
																	
																	
																<table border=2 cellpadding=10><tr><td>	
																	<b>
																	<font face="Calibri" size="4" color="#008000">
																	PERIODE 
																	JAN-JUN 2016</font></b><font color="#000000"><br>
												<b>Point yang telah anda 
																	dapatkan 
																	adalah 
																	sebanyak </b>
												</font><b>
												<% rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2016-01' and tgl < '2016-06-32' order by id ",conn
												if rs.bof then
												point = 0
												else
												point = rs("sum(point)")
												end if
												rs.close
												
												rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2016-01-32' and tgl < '2016-07-25' order by id ",conn
												if rs.bof then
												pointo = 0
												else
												pointo = rs("sum(point)")
												end if
												rs.close
												
												if isnull(pointo) = true then
												pointo = 0
												end if
												ponta = point - pointo
												%>
												<font color="#FF0000"><%=point%> </font>
																	,</b><font color="#000000"><b> 
																	dan telah 
																	klaim <font color="#FF0000"><%=pointo%> </font>
																	, jadi sisa <font color="#FF0000"><%=ponta%> </font> 
																	Point RP</b><br>
												<b>History Sponsorisasi anda :
												<% rs.Open "SELECT direk,paket,point,tgl FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2016-01' and tgl < '2016-06-32' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Paket</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("paket"))%> - <%=ucase(rs("direk"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>
<b>History Klaim anda :
												<% rs.Open "SELECT * FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2016-01-32' and tgl < '2016-07-25' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Keterangan</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Sudah Cetak</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
idklaim = rs("id")
ket = rs("ket")
ambil = rs("ambil")

if ambil = "F" then
warna = "white"
else
warna = "yellow"
end if

%>			
<tr bgcolor="<%=warna%>"><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("hadiah"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td><td align="center"><form name="lihatdc2<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idklaim" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<font color="#000000"><b>pengambilan : </b></font>
				<input name="ket" value="<%=ket%>"style="font-size: 8pt; font-family: Verdana; border: 1px solid #808080; font-weight:700" size="40"><font color="#000000"><b>
				</b></font> 
				<input type="submit" name="btpoint" value="Update" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form></td><td align="center"><font color="#000000"><b><%if warna = "white" then%></b></font><form name="updatrdc2<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idambil" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<input type="submit" name="btpoint" value="Sudah" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form><font color="#000000"><b><%end if%></b></font></td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table><b>Yang sudah di Ambil :
<% rs.Open "SELECT * FROM healthwealthint_com_hwi.fx_ge_det where nopo like '%"&ktana&"%' ",conn,3,3 %>
												</b>
																	</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>No invoice</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Jumlah</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("nopo"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("kode"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("jumlah"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>

<br>&nbsp;</font><p></td></tr></table>
					<table border=2 cellpadding=10><tr><td>	
																	<b>
																	<font face="Calibri" size="4" color="#008000">
																	PERIODE 
																	JUL-DES 2016</font></b><font color="#000000"><br>
												<b>Point yang telah anda 
																	dapatkan 
																	adalah 
																	sebanyak </b>
												</font><b>
												<% rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2016-07' and tgl < '2016-12-32' order by id ",conn
												if rs.bof then
												point = 0
												else
												point = rs("sum(point)")
												end if
												rs.close
												
												rs.Open "SELECT sum(point) FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2016-07-25' and tgl < '2017-01-06' order by id ",conn
												if rs.bof then
												pointo = 0
												else
												pointo = rs("sum(point)")
												end if
												rs.close
												
												if isnull(pointo) = true then
												pointo = 0
												end if
												ponta = point - pointo
												%>
												<font color="#FF0000"><%=point%> </font>
																	,</b><font color="#000000"><b> 
																	dan telah 
																	klaim <font color="#FF0000"><%=pointo%> </font>
																	, jadi sisa <font color="#FF0000"><%=ponta%> </font> 
																	Point RP</b><br>
												<b>History Sponsorisasi anda :
												<% rs.Open "SELECT direk,paket,point,tgl FROM healthwealthint_com_hwi.tab_rp_sponsor where kta like '"&ktana&"' and tgl > '2016-07' and tgl < '2016-12-32' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Paket</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("paket"))%> - <%=ucase(rs("direk"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>
<b>History Klaim anda :
												<% rs.Open "SELECT * FROM healthwealthint_com_hwi.tab_rp_sponsor_release where kta like '"&ktana&"' and tgl > '2016-07-25' and tgl < '2017-01-06' order by id ",conn,3,3 %>
												</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Tanggal</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>Hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Poin</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Keterangan</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Sudah Cetak</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
idklaim = rs("id")
ket = rs("ket")
ambil = rs("ambil")

if ambil = "F" then
warna = "white"
else
warna = "yellow"
end if

%>			
<tr bgcolor="<%=warna%>"><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("tgl"))%></b></font></td><td align="center" >
<font color="#000000"><b>
<%=ucase(rs("hadiah"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("point"))%> </b></font> </td><td align="center"><form name="lihatdc<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idklaim" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<font color="#000000"><b>pengambilan : </b></font>
				<input name="ket" value="<%=ket%>"style="font-size: 8pt; font-family: Verdana; border: 1px solid #808080; font-weight:700" size="40"><font color="#000000"><b>
				</b></font> 
				<input type="submit" name="btpoint" value="Update" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form></td><td align="center"><font color="#000000"><b><%if warna = "white" then%></b></font><form name="updatrdc<%=idklaim%>" method="post" action="" onsubmit="return formCheck1(this)">
				<input type="hidden" name="idambil" value="<%=idklaim%>" style="font-weight: 700" >
				<input type="hidden" name="ktana" value="<%=rs("kta")%>" style="font-weight: 700" >
				<input type="submit" name="btpoint" value="Sudah" style="font-size: 8pt; font-family: Verdana; font-weight:700"><font color="#000000"><b>
				</b></font>
				</form><font color="#000000"><b><%end if%></b></font></td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table><b>Yang sudah di Ambil :
<% rs.Open "SELECT * FROM healthwealthint_com_hwi.fx_ge_det where nopo like '%"&ktana&"%' ",conn,3,3 %>
												</b>
																	</b>
<table border="1" cellpadding="0" cellspacing="0" width="80%" ><tr bgcolor="#99FF00" >
	<td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>No invoice</b></font></td><td align="center" bgcolor="#C0C0C0">
		<font color="#000000"><b>hadiah</b></font></td><td align="center" bgcolor="#C0C0C0">
	<font color="#000000"><b>Jumlah</b></font></td></tr>
<%
lumpat = 2000
for b = 1 to lumpat
if rs.eof = True then 
	exit for 
	b = b + 1
end if
%>			
<tr><td align="center">
<font color="#000000"><b>
<%=ucase(rs("nopo"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("kode"))%></b></font></td><td align="center">
<font color="#000000"><b>
<%=ucase(rs("jumlah"))%> </b></font> </td></tr>
<%  
rs.movenext
if b > lumpat then
exit for
end if
next
rs.close 
%>	
</table>

<br>&nbsp;</font><p></td></tr></table>												
																	&nbsp;</td>
																</tr>
																</table>
																
																<%
																end if
																%>
																
												</td>
											</tr>
										</table>
									</div>
									</td>
								</tr>
								</table>
							</td>
						</tr>
					</table>
					</td>
				</tr>
				<tr>
					<td valign="top">&nbsp;</td>
				</tr>
			</table>
		</div>
		</td>
	</tr>
</table>
<%
set rs=nothing
Conn.Close
set conn=nothing

set rsALL=nothing
ConnALL.Close
set connALL=nothing
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td><!--#Include File=pg_footer.asp--></td>
	</tr>
</table>

</body>

</html>