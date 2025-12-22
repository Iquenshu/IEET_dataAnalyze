<!-- #include file="../funs.txt" -->
<%
server.scripttimeout=6000   

disc= array("0-9","10-19","20-29","30-39","40-49","50-59","60-69","70-79","80-89","90-100")
redim rate(9), rate1(9),prate(9), prate1(9)
redim avgn(300),avgs0(300),avgs1(300)
syys= trim(request("syy"))
if syys = "" then
	syy=107
else
    syy=Cint(syys)	
end if		
dbn="../eestudents.mdb"
set cn=db_connection(dbn)
dbn="../teaching/courses.mdb"
set cn1=db_connection(dbn)
'識別碼	學年度	學期	系所代碼	系所名稱	課號	課程名稱	學號	姓名	成績

'SQL="select 課程名稱 from 電機系學生成績 where 課號 like 'B%' and 課程名稱='計算機概論'  group by 課程名稱 order by 課程名稱"
'SQL="select 課程名稱 from 電機系學生成績 where 課號 like 'B%' and 學年度= '"&Cstr(syy)&"' group by 課程名稱"

%>
<html>

<head>
<title>電機系各課程學生成績分佈</title>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
</head>

<body>

<p align="center"><font size="6">電機系各課程學生成績分佈[<%=syy%>]</font></p>
<center>
<%
SQL="select 學期,科目名稱,年級 from 學期課程 where  學年= '"&Cstr(syy)&"' order by 學期, 年級, 科目名稱 "
'SQL="select 學期,科目名稱,年級 from 學期課程 where  學年= '"&Cstr(syy)&"' and 科目名稱 in ('電子學（三）_乙','電子學（三）_甲') order by 學期, 年級, 科目名稱 "
set ccsrs2=open_recordset(cn1,SQL,3,3)

k=0
while not ccsrs2.eof 
csnA=split(ccsrs2("科目名稱"),"_")

if trim(csnA(0)) <> "" then

SQL="select 課程名稱,課號 from 電機系學生成績 where 課程名稱='"&trim(csnA(0))&"'  and 學年度= '"&Cstr(syy)&"' group by 課程名稱,課號"
'response.write "["&SQL&"]<br>"
set csrs=open_recordset(cn,SQL,3,3)

if csrs.recordcount > 0 then
if ubound(csnA) = 1 then
   cno=3
   clsshn= trim(csnA(1))
   if clsshn = "甲" then
       cno=1     
   end if    
   if clsshn = "乙" then
       cno=2
   end if    
else           
   cno=0
   clsshn= "" 
end if  
select case trim(ccsrs2("年級"))
    case "大一"
        ccno="1"
    case "大二"
        ccno= "2"
    case "大三"
        ccno= "3"
    case "大四"
        ccno= "4"
    case else
        ccno= "_"
end select    
    
csno=csrs("課號")       
if cno=2 and syy >= 101 then
	csno=csno&"A"
else
   if syy < 101 then
       csno="B301"&ccno&"2"
   end if    	
end if    
 
csno1=csrs("課號")
if cno=2 and syy >= 101 then
	csno1=csno1&"A"
else
   if syy < 101 then
       csno1="B301"&ccno&"2"
   end if    	
end if    
        
%>
<br><br>
<p align="center"><font size="4">[<%=csrs("課程名稱")&"("&ccsrs2("學期")&","&ccsrs2("年級")&")"%><%=clsshn%>]</font><br>
<%
avgn(k)=csrs("課程名稱")&"("&ccsrs2("學期")&","&ccsrs2("年級")&")"&clsshn

sc_satistic csrs("課程名稱"),Cstr(syy-1),Cstr(cno),"_",disc,rate,sct1,csno1&" "
sum1=percent(rate,prate)
if sum1 > 0 then
	av1=fdigit(sct1/sum1)
else
    av1=0
end if    
avgs0(k)=av1
	
sc_satistic csrs("課程名稱"),Cstr(syy),Cstr(cno),ccno,disc,rate1,sct2,csno&" "
sum2=percent(rate1,prate1)
if sum2 > 0 then
	av2=fdigit(sct2/sum2)
else
    av2=0
end if    	
avgs1(k)=av2
%>
<font size="3"><%=Cstr(syy-1)%> 學年度平均[<%=av1%> ] 學生人數[<%=sum1%>]</font><br>
<font size="3"><%=Cstr(syy)%> 學年度平均[<%=av2%> ] 學生人數[<%=sum2%>]</font></p>

<%
drawG Cstr(syy-1)&" 學年度",Cstr(syy)&" 學年度",disc,rate,prate,rate1,prate1
k=k+1
'if k mod 2 = 0 then
%>
 <br style="page-break-after:always">
<%
'end if
csrs.close
%>

<%
end if
end if

ccsrs2.movenext
wend
cn.close
cn1.close
%>
<br style="page-break-after:always">
<table border="1" width="650">
	<tr>
		<td width="356">課程名稱</td>
		<td width="145">[<%=syy-1%>]學年度</td>
		<td width="134">[<%=syy%>]學年度</td>
	</tr>
<%
for i=0 to k-1
%>	
	<tr>
		<td width="356"><%=avgn(i)%></td>
		<td width="145"><%=avgs0(i)%></td>
		<td width="134"><%=avgs1(i)%></td>
	</tr>
<%
next
%>	
</table>
<p>
</p>
</center>

</body>

</html>

<%
function percent (rate,prate)
dim i,sum
sum=0
for i= 0 to ubound(rate)
sum=sum+rate(i)
prate(i) = 0
next

if sum <> 0 then
for i= 0 to ubound(rate)
prate(i) = (rate(i) * 100) \ sum
next
end if
percent=sum
end function

sub sc_satistic(csno,syy,cno,ccno,disc,rate,sct,ccnos)
dim dbn,sql,cn,rs,ta,i,sctt,k
dbn="../eestudents.mdb"
set cn=db_connection(dbn)
'識別碼	學年度	學期	系所代碼	系所名稱	課號	課程名稱	學號	姓名	成績
sctt=0
for i=0 to ubound(disc)
    ta=split(disc(i),"-")
    if cno = 0 then
		SQL="select 成績 from 電機系學生成績 where 課程名稱='"& trim(csno) &"' and 學年度= '"& syy &"' and ( 成績 >= " & trim(ta(0)) & " and 成績 <= "& trim(ta(1))& ")"
	else
	    if cno < 3 then
	        if syy <= 100 then
		        SQL="select 成績 from 電機系學生成績 where 課程名稱='"& trim(csno) &"' and 課號 like 'B301"&ccno&trim( cno) &"%' and 學年度= '"& syy &"' and ( 成績 >= " & trim(ta(0)) & " and 成績 <= "& trim(ta(1))& ")"
		    else
		        SQL="select 成績 from 電機系學生成績 where 課程名稱='"& trim(csno) &"' and 課號='"&trim(ccnos)&"' and 學年度= '"& syy &"' and ( 成績 >= " & trim(ta(0)) & " and 成績 <= "& trim(ta(1))& ")"
		    end if    
		else    
		    SQL="select 成績 from 電機系學生成績 where 課程名稱='"& trim(csno) &"' and 課號='"&trim(ccnos)&"' and 學年度= '"& syy &"' and ( 成績 >= " & trim(ta(0)) & " and 成績 <= "& trim(ta(1))& ")"
		end if 
  
	end if		
'response.write "["&cno&"]["&SQL&"]<br>"
	set rs=open_recordset(cn,SQL,3,3)
	rate(i)=rs.recordcount
	if rs.recordcount > 0 then
	    for k= 0 to rs.recordcount-1
	        sctt=sctt+rs("成績")
	        rs.movenext
	    next    
	    rs.close
	end if
next	    
cn.close
sct=sctt
end sub

function get_course_name(csno)
dim dbn,sql,cn,rs
dbn="../eestudents.mdb"
set cn=db_connection(dbn)
'識別碼	學年度	學期	系所代碼	系所名稱	課號	課程名稱	學號	姓名	成績

SQL="select 課程名稱 from 電機系學生成績 where 課號='"&trim(csno)&"'"
set rs=open_recordset(cn,SQL,3,3)
if rs.bof and rs.eof then
    get_course_name=""
else    
    get_course_name=rs("課程名稱")
    rs.close
end if
cn.close
end function

sub drawG(dn1,dn2,disc,rate,prate,rate1,prate1)
dim i
%>  
<table border="0" width="623" id="table1" cellspacing="0" cellpadding="0" >

<tr>
		<td  width="100" align=right><font size="3"><%=dn1%></font></td>
		<td width="523" align="left" >
		<img border="0" src="../images/greenbar.JPG" width="30%" height="10">		
		</td>

</tr>
<tr>
		<td  width="100" align=right><font size="3"><%=dn2%></font></td>
		<td width="523" align="left" >
		<img border="0" src="../images/Ybar.JPG" width="30%" height="10"><br>		
		</td>

</tr>
</table>
<hr width="50%">
<table border="0"width="650" id="table1" cellspacing="0" cellpadding="0" >
<%
for i=0 to ubound(disc)
%>
	<tr>
		<td width="100" align=right><font size="3"><%=disc(i)%></font></td>
		<td width="540">
		<table border="0" cellspacing="0" cellpadding="0" >
			<tr>			
				<td   align=right width="10" style="border-style: solid; border-width: 0px; padding-left: 2px; padding-right: 2px; padding-top: 0px; padding-bottom: 0px"><span style="font-size: 8.0pt; font-family: Arial; color: black"><%=dg2(rate(i))%> </span> </td>
				<td   align=right width="20" style="border-style: solid; border-width: 0px; padding-left: 2px; padding-right: 2px; padding-top: 0px; padding-bottom: 0px"><span style="font-size: 8.0pt; font-family: Arial; color: black"><%=dg2(prate(i))%>
				% </span> </td>
				<td width="500" align="left" ><span style="font-size: 8.0pt; font-family: Arial; color: black">
<%
wpd= (prate(i) * 500) / 100 +0.5
wpds= Cstr(wpd)
taa=split(wpds,".")
wpds=taa(0)
%>
				<img border="0" src="../images/greenbar.JPG" width="<%=wpds%>" height="10">
				</span> 
				</td>
			</tr>
			<tr>
				<td  align=right width="10" style="border-style: solid; border-width: 0px; padding-left: 2px; padding-right: 2px; padding-top: 0px; padding-bottom: 0px"><span style="font-size: 8.0pt; font-family: Arial; color: black"><%=dg2(rate1(i))%> </span> </td>
				<td  align=right width="20" style="border-style: solid; border-width: 0px; padding-left: 2px; padding-right: 2px; padding-top: 0px; padding-bottom: 0px"><span style="font-size: 8.0pt; font-family: Arial; color: black"><%=dg2(prate1(i))%>
				% </span> </td>
				<td width="500" align="left"><span style="font-size: 8.0pt; font-family: Arial; color: black">
<%
wpd= (prate1(i) * 500) / 100+0.5
wpds= Cstr(wpd)
taa=split(wpds,".")
wpds=taa(0)
%>
				
				<img border="0" src="../images/Ybar.JPG" width="<%=wpds%>" height="10"> 			 
				</span> </td>				
			</tr>
		</table>
		</td>
	</tr>
	<tr>
	<td colspan=2><span style="font-size: 2.0pt; font-family: 標楷體; color: black"> 
	&nbsp;&nbsp; </span></td>
	</tr>
<%next%>	

</table>
<hr width="50%">
<%
end sub
function dg2(ff)
dim ffs,ta
    ffs=Cstr(ff)
    if len(ffs) = 1 then
        dg2="0"&ffs
    else
        dg2=ffs
    end if        
end function

function fdigit(ff)
dim ffs,ta
    ffs=Cstr(ff)
    ta=split(ffs,".")
    if ubound(ta) >=1 then
    	fdigit=ta(0)&"."&Mid(ta(1),1,2)
    else
    	fdigit=ta(0)
    end if	
end function

function StudentCount(yys, item)
dim dbn,sql,cn,rs
dbn="../eestudents.mdb"
SQL="select 學校名稱 from 電機系入學前學校 where 學校名稱 like '%"&item&"%' and 入學年='"&trim(yys)&"'"
set cn=db_connection(dbn)
set rs=open_recordset(cn,SQL,3,3)
n1=rs.recordcount
rs.close
cn.close

StudentCount= n1

end function

function GetFildWidth(tbln,rs,n)
dim i,k
dim s,wrs,sql,ss,wsum
Redim wlen(rs.recordcount)
s=" "
'tbln	fieldn	flen

k=0
wsum=0
for i= 1 to n-1
    sql="select flen from tblmgr where tbln='"&tbln&"'  and fieldn='"&trim(rs(i).name)&"'"
	set wrs=open_recordset(cn,SQL,3,3)
	if not (wrs.bof and wrs.eof) then
	    wlen(k)=wrs("flen")
	    wsum=wsum+wlen(k)
	    k=k+1
	    wrs.close
	end if
next

if wsum <> 0 then
    for i= 0 to k-1
    
        ss= Cstr((wlen(i)\wsum)*100) 
        if trim(s) = "" then
            s="3,"&ss
        else
            s=s&","&ss
        end if
    next   
end if             

GetFildWidth=s

end function

%>