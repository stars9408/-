<!--#include virtual="connect.asp" -->
<!-- #include virtual="Lib/Asp/Func/Customer_Union.Func.asp" -->

<!--#include virtual="/include/default.asp" -->

<script type="text/javascript" src="/jscript/callme.js"></script>
<script type="text/javascript" src="/jscript/consultation.js"></script>
<%
Server.ScriptTimeout = 99999999
response.Expires=0
auth_consultation = session("auth_consultation")
SQLTyear = " SELECT  * FROM group_list WHERE  (CODE_NO= '40')  and emp_id='"&session("id")&"'"
Set rsTyear = db.Execute(SQLTyear)
IF NOT rsTyear.EOF THEN
	session("auth_consultation") ="999"
	auth_consultation="999"
END IF
	response.Buffer=true
	Call Loading_Msg()
onchangechk = request("onchangechk")


IF request("matchQuarter")="" THEN
matchQuarter= "1"
ELSE
matchQuarter=request("matchQuarter")
END If

If request("matchQuarter") = "" Or request("matchQuarter") = "1" Then
	Month1="01"
	Month2="02"
	Month3="03"
ElseIf request("matchQuarter") = "2" Then
	Month1="04"
	Month2="05"
	Month3="06"
ElseIf request("matchQuarter") = "3" Then
	Month1="07"
	Month2="08"
	Month3="09"
ElseIf request("matchQuarter") = "4" Then
	Month1="10"
	Month2="11"
	Month3="12"
END If

IF request("tyear1")="" THEN
	tyear1 = year(date)
	tyear2 = tyear1
ELSE
	tyear1 = request("tyear1")
	tyear2 = tyear1
END If
	matchday1 = "01"
	matchday2 = "31"

'�˻���
	'sqlM1 ="SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month1 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month1 &"-"& matchday2 &"'  order by contract_date_tax"
	
	'sqlM2 ="SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month2 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month2 &"-"& matchday2 &"'  "

	'sqlM3 ="SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month3 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month3 &"-"& matchday2 &"'  "

	sqlinfrait ="SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month1 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month3 &"-"& matchday2 &"'  and doc_kind in('9','10','11','13','14', '17', '18', '19', '20') "

	sqlkrinfra ="SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month1 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month3 &"-"& matchday2 &"'  and doc_kind not in('9','10','11','13','14', '17', '18', '19', '20') "
	
	'Set rs  = db.Execute(sqlM1)
	Set rs4 = db.Execute(sqlinfrait)
	Set rs5 = db.Execute(sqlkrinfra)
	'response.write sqlM1
	
%>
<script language="JavaScript" type="text/JavaScript">
<!--
function sel_name()
{
  var matchMonth = list.matchMonth.value;
  var tyear = list.tyear.value;
  location.href="profitCalculator.asp?matchMonth="+matchMonth+"&tyear="+tyear+"";
}

function openWindows1(vStr) {
  window.open (vStr, "win44", 'width=550,height=550,left=100,top=100,resizable=1,scrollbars=yes, menubar=yes');
}

function setonchangechk(fm){
	fm.onchangechk.value = "T";
	fm.submit();
}
function printTime() {
		var clock = document.getElementById("clock");
		var now = new Date();
		clock.innerHTML = now.getFullYear() + "." + (now.getMonth()+1) + "." + now.getDate() ;
		setTimeout("printTime()", 1000); 
	} 
window.onload = function() {
	printTime();
	}; 
-->
</script>


<script type="text/javascript" src="/jcript/callme.js"></script>

<div class="conSecWide" style="width:4800px;">
	<h2><i class="fa fa-folder-open-o"></i>��������<span class="subTitle">���Ͱ�꼭</span></h2>
	<div class="tb_topWrap">
		<form name="list_form" method="post" action="profitCalculator.asp">
			<input type="hidden" name="onchangechk" value="">
			<input type="hidden" name="userdeptchk" value="">
			<span class="infoWrap_left">
				<select name='tyear1' onChange="setonchangechk(this.form)" class="form_basic">
				<%
					SQL = "Select tyear From ����Ī group by tyear order by tyear desc"
					SET Rs_tyear = db.execute(sql)

						DO UNTIL Rs_tyear.eof 
				%>
						<option value="<%=Trim(Rs_tyear("tyear"))%>" <%if Trim(tyear1)=Trim(Rs_tyear("tyear")) then%>selected<%end if%>><%=Rs_tyear("tyear")%></option>
				<%
						Rs_tyear.movenext
						LOOP
				%>
				</select>
				<label>��</label>

				<select name='matchQuarter'  onchange="setonchangechk(this.form)" class="form_basic">
				<%FOR j=1 To 4%>				
					<option value="<%=j%>" <%if trim(cint(matchQuarter))=trim(j) then%>selected<%end if%>><%=j%></option>         
				<% NEXT %>
				</select>
				<label>Quarter</label>
				<input type="submit" class="fa btn_t02 btnGray" value="&#xf007 �˻�">
				<a href="profitCalculator_Excel.asp" class="btn effect01" target="_blank"><span>���� �ٿ�ε�</span></a>
				
			</span>
		</form>
	<div style="clear:both;float:left;width:100%;font-size:14px;text-align:left;padding:20px 0;">
		<strong>��&nbsp;���� �б� ��� ����&nbsp;Q<%=matchQuarter%></strong>&nbsp;<b><<span id="clock"></b></span>>
	</div>

	<div style = "clear:both;"></div>
	<div style="text-align:left;"><b>�� ���� ���� ��� ����</b></div>
	<table border="0" cellpadding="0" cellspacing="0" class="tb_basic" style="width:1000px">
		<tr>
			<th colspan="9"><div>�� ���� ���� ��� ����</div><div style="text-align:right;"> - Q<%=matchQuarter%> -</div></th>
		</tr>
		<tr> 
			<th rowspan="3">Q<%=matchQuarter%></th>
			<th colspan="2">M1 : <%=Month1%>��</th>
			<th colspan="2">M2 : <%=Month2%>��</th>
			<th colspan="2">M3 : <%=Month3%>��</th>
			<th colspan="2">�հ�</th>
		</tr>
		<tr> 
			<td style="text-align:center;">����</td>
			<td style="text-align:center;">������</td>
			<td style="text-align:center;">����</td>
			<td style="text-align:center;">������</td>
			<td style="text-align:center;">����</td>
			<td style="text-align:center;">������</td>
			<td style="text-align:center;">����</td>
			<td style="text-align:center;">������</td>
		</tr>
		<tr>
		<!--���� ����-->
<% 
	monthnow = month1
	months = 0
	currentmonth = month1
	Do Until months = 3
	amount_money_SUM_TOT = 0
	product_amount_SUM_1 = 0
	sell_pay_amount_SUM_1 = 0
	exe_Rs_pay_amount_SUM_1 = 0
	exe_Rs_1_pay_amount_SUM_1 = 0
		sqlM1 ="SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& monthnow &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& monthnow &"-"& matchday2 &"'  order by contract_date_tax"
		Set rs  = db.Execute(sqlM1)
		'Response.write sqlM1
	'�� ���
	Do until rs.eof   
			id = rs("id")
			user_name = rs("user_name")
			contract_date_tax = rs("contract_date_tax")
			amount_money = rs("amount_money")

			' ��ǰ����,���Ⱥ��
			'----------------------------------------------------------------------------
					sql = "select amount, account, account_des from ������� where id = '" &id& "' "
					Set pro_Rs = db.Execute(sql)	
					
					' ��ǰ����, ��Ÿ���, ���Ⱥ�� ���
					Do until pro_Rs.eof
						amount=pro_Rs("amount") '�ݾ�
						account=Trim(pro_Rs("account"))'����
						account_des=Trim(pro_Rs("account_des"))'����

						IF account = "��ǰ����" Then
						product_amount_SUM_1 = product_amount_SUM_1 + amount
						ElseIf account = "���Ⱥ��" Then
						product_amount_SUM_1 = product_amount_SUM_1 + amount
						ELSEIF account_des = "��Ÿ" THEN 
						product_amount_SUM_1 = product_amount_SUM_1 + amount
						End If
					pro_Rs.MoveNext
					Loop

					'���ְ��߿��� ���
					devmoneysql = "select tamount, services from ������ް�ȹ where id = '" &id& "' "
					Set pay_Rs = db.Execute(devmoneysql)	

					If Not pay_Rs.eof Then
						amount=pay_Rs("tamount") '�ݾ�
						services=Trim(pay_Rs("services")) '����

						IF services = "1" THEN 
						sell_pay_amount_SUM_1 = sell_pay_amount_SUM_1 + amount
						ELSEIF services = "2" THEN 
						sell_pay_amount_SUM_1 = sell_pay_amount_SUM_1 + amount
						END If
					End If
					
					' ���࿹��
			'----------------------------------------------------------------------------
					sql = "select sum(amount) as amount from ���࿹�� where id = '" &id& "'  group by id "
					Set exe_Rs = db.Execute(sql)	

					If Not exe_Rs.eof Then
					amount=exe_Rs("amount") '�ݾ�
					exe_Rs_pay_amount_SUM_1 = exe_Rs_pay_amount_SUM_1 + amount
					End If

					
					sql = "select sum(amount) as amount from �Ҹ𼺰�� where id = '" &id& "'   group by id "
					Set exe_Rs_1 = db.Execute(sql)	

					If Not exe_Rs_1.eof Then
					amount=exe_Rs_1("amount") '�ݾ�
					exe_Rs_1_pay_amount_SUM_1 = exe_Rs_1_pay_amount_SUM_1 + amount
					End If

					exe_Rs_pay_amount_SUM = exe_Rs_pay_amount_SUM_1 + exe_Rs_1_pay_amount_SUM_1
			'�޺� ����, ������
			amount_money_SUM_TOT = amount_money_SUM_TOT + amount_money
			amount_money_total = amount_money_SUM_TOT - product_amount_SUM_1 - sell_pay_amount_SUM_1 - exe_Rs_pay_amount_SUM

			Rs.Movenext
			Loop

			'���� ����, ������ �հ�
			Total_amount_money_SUM = Total_amount_money_SUM + amount_money_SUM_TOT
			Total_earn_money_Sum   = Total_earn_money_Sum + amount_money_total + amount_money_total2 + amount_money_total3

			If currentmonth = month1 Then
				monthnow = month2
				currentmonth = month2
			ElseIf currentmonth = month2 Then
				monthnow = month3
			End If
			months = months + 1
			%>

			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_TOT,0)%>��</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_total,0)%>��</b></td>	

			<%
			Loop	
			%>

			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(Total_amount_money_SUM,0)%>��</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(Total_earn_money_Sum,0)%>��</b></td>
		</tr>
		</table>
		
		<%

		'������ ������� ���
		Do until rs4.eof   
			id = rs4("id")
			user_name = rs4("user_name")
			contract_date_tax = rs4("contract_date_tax")
			amount_money = rs4("amount_money")
			' ��ǰ����,���Ⱥ��
			'----------------------------------------------------------------------------
					sql = "select amount, account, account_des from ������� where id = '" &id& "' "
					Set pro_Rs = db.Execute(sql)	
					
					' ��ǰ����, ��Ÿ���, ���Ⱥ�� ���
					Do until pro_Rs.eof
						amount=pro_Rs("amount") '�ݾ�
						account=Trim(pro_Rs("account"))'����
						account_des=Trim(pro_Rs("account_des"))'����

						IF account = "��ǰ����" Then
						product_amount_SUM_4 = product_amount_SUM_4 + amount
						ElseIf account = "���Ⱥ��" Then
						product_amount_SUM_4 = product_amount_SUM_4 + amount
						ELSEIF account_des = "��Ÿ" THEN 
						product_amount_SUM_4 = product_amount_SUM_4 + amount
						End If
					pro_Rs.MoveNext
					Loop

					'���ְ��߿��� ���
					devmoneysql = "select tamount, services from ������ް�ȹ where id = '" &id& "' "
					Set pay_Rs = db.Execute(devmoneysql)	

					If Not pay_Rs.eof Then
						amount=pay_Rs("tamount") '�ݾ�
						services=Trim(pay_Rs("services")) '����

						IF services = "1" THEN 
						sell_pay_amount_SUM_4 = sell_pay_amount_SUM_4 + amount
						ELSEIF services = "2" THEN 
						sell_pay_amount_SUM_4 = sell_pay_amount_SUM_4 + amount
						END If
					End If
					
					' ���࿹��
			'----------------------------------------------------------------------------
					sql = "select sum(amount) as amount from ���࿹�� where id = '" &id& "'  group by id "
					Set exe_Rs = db.Execute(sql)	

					If Not exe_Rs.eof Then
					amount=exe_Rs("amount") '�ݾ�
					exe_Rs_pay_amount_SUM_4 = exe_Rs_pay_amount_SUM_4 + amount
					End If

					
					sql = "select sum(amount) as amount from �Ҹ𼺰�� where id = '" &id& "'   group by id "
					Set exe_Rs_1 = db.Execute(sql)	

					If Not exe_Rs_1.eof Then
					amount=exe_Rs_1("amount") '�ݾ�
					exe_Rs_1_pay_amount_SUM_4 = exe_Rs_1_pay_amount_SUM_4 + amount
					End If

					exe_Rs_pay_amount_SUM4 = exe_Rs_pay_amount_SUM_4 + exe_Rs_1_pay_amount_SUM_4

			amount_money_SUM_TOT4 = amount_money_SUM_TOT4 + amount_money
			amount_money_total4 = amount_money_SUM_TOT4 - product_amount_SUM_4 - sell_pay_amount_SUM_4 - exe_Rs_pay_amount_SUM4


		  Rs4.Movenext
		  Loop

		  '�ѱ������� ���
		Do until rs5.eof   
			id = rs5("id")
			user_name = rs5("user_name")
			contract_date_tax = rs5("contract_date_tax")
			amount_money = rs5("amount_money")
			' ��ǰ����,���Ⱥ��
			'----------------------------------------------------------------------------
					sql = "select amount, account, account_des from ������� where id = '" &id& "' "
					Set pro_Rs = db.Execute(sql)	
					
					' ��ǰ����, ��Ÿ���, ���Ⱥ�� ���
					Do until pro_Rs.eof
						amount=pro_Rs("amount") '�ݾ�
						account=Trim(pro_Rs("account"))'����
						account_des=Trim(pro_Rs("account_des"))'����

						IF account = "��ǰ����" Then
						product_amount_SUM_5 = product_amount_SUM_5 + amount
						ElseIf account = "���Ⱥ��" Then
						product_amount_SUM_5 = product_amount_SUM_5 + amount
						ELSEIF account_des = "��Ÿ" THEN 
						product_amount_SUM_5 = product_amount_SUM_5 + amount
						End If
					pro_Rs.MoveNext
					Loop

					'���ְ��߿��� ���
					devmoneysql = "select tamount, services from ������ް�ȹ where id = '" &id& "' "
					Set pay_Rs = db.Execute(devmoneysql)	

					If Not pay_Rs.eof Then
						amount=pay_Rs("tamount") '�ݾ�
						services=Trim(pay_Rs("services")) '����

						IF services = "1" THEN 
						sell_pay_amount_SUM_5 = sell_pay_amount_SUM_5 + amount
						ELSEIF services = "2" THEN 
						sell_pay_amount_SUM_5 = sell_pay_amount_SUM_5 + amount
						END If
					End If
					
					' ���࿹��
			'----------------------------------------------------------------------------
					sql = "select sum(amount) as amount from ���࿹�� where id = '" &id& "'  group by id "
					Set exe_Rs = db.Execute(sql)	

					If Not exe_Rs.eof Then
					amount=exe_Rs("amount") '�ݾ�
					exe_Rs_pay_amount_SUM_5 = exe_Rs_pay_amount_SUM_5 + amount
					End If

					
					sql = "select sum(amount) as amount from �Ҹ𼺰�� where id = '" &id& "'   group by id "
					Set exe_Rs_1 = db.Execute(sql)	

					If Not exe_Rs_1.eof Then
					amount=exe_Rs_1("amount") '�ݾ�
					exe_Rs_1_pay_amount_SUM_5 = exe_Rs_1_pay_amount_SUM_5 + amount
					End If

					exe_Rs_pay_amount_SUM5 = exe_Rs_pay_amount_SUM_5 + exe_Rs_1_pay_amount_SUM_5

			amount_money_SUM_TOT5 = amount_money_SUM_TOT5 + amount_money
			amount_money_total5 = amount_money_SUM_TOT5 - product_amount_SUM_5 - sell_pay_amount_SUM_5 - exe_Rs_pay_amount_SUM5


		  Rs5.Movenext
		  Loop
	%>

	<div style = "clear:both;"></div>
	<div style="text-align:left;"><b>�� (��)�ѱ������� & (��)������������� ���� ����</b></div>
	<table border="0" cellpadding="0" cellspacing="0" class="tb_basic" style="width:1000px">
		<tr>
			<th colspan="7"><div>�� (��)�ѱ������� & (��)������������� ���� ����</div><div style="text-align:right;"> - Q<%=matchQuarter%> -</div></th>
		</tr>
		<tr> 
			<th rowspan="2">����</th>
			<th colspan="3">����</th>
			<th colspan="3">��������</th>
		</tr>
		<tr> 
			<td style="text-align:center;">�ѱ�������</td>
			<td style="text-align:center;">�������</td>
			<td style="text-align:center;">�հ�</td>
			<td style="text-align:center;">�ѱ�������</td>
			<td style="text-align:center;">�������</td>
			<td style="text-align:center;">�հ�</td>
		</tr>
		<tr> 
			<th rowspan="1">Q<%=matchQuarter%></th>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_TOT5,0)%>��</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_TOT4,0)%>��</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_TOT5 + amount_money_SUM_TOT4,0)%>��</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_total5,0)%>��</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_total4,0)%>��</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_total5 + amount_money_total4,0)%>��</b></td>
		</tr>
	</table>
	<br><br>
	<div style = "clear:both;"></div>
	<div style="text-align:left;"><b>�� ����  �б� �������</b></div>
	<table width="100%" border="0" cellpadding="0" cellspacing="0" class="tb_basic" >
		<tr>
			<th><div>�� ����  �б� �������</div><div style="text-align:right;"> - Q<%=matchQuarter%> -</div></th>
		</tr>
	</table>


<%	
	sqlPersonSum = "SELECT Top 1 count(a.user_dept) as max FROM (SELECT userlist.emp_id, userlist.user_name, userlist.user_dept FROM userlist INNER JOIN consultation ON userlist.emp_id = consultation.emp_id WHERE  (consultation.approval_flag = '1' or consultation.approval_flag = '3') and (left(consultation.contract_date_tax,4) >= '"&tyear1&"' AND left(consultation.contract_date_tax,4) <= '"&tyear2&"') GROUP BY userlist.emp_id, userlist.user_name, userlist.user_dept ) a GROUP BY a.user_dept order by max desc "
		Set personNum = db.Execute(sqlPersonSum)
			If Not personNum.eof Then
				MaxPersonNum = personNum("max")
			End If

	monthnow = month1
	months = 0
	currentmonth = month1
	Do Until months = 3
	'���κ� ����


	sqlPerson = "SELECT userlist.emp_id, userlist.user_name, userlist.user_dept FROM userlist INNER JOIN consultation ON userlist.emp_id = consultation.emp_id WHERE  (consultation.approval_flag = '1' or consultation.approval_flag = '3') and (left(consultation.contract_date_tax,4) >= '"&tyear1&"' AND left(consultation.contract_date_tax,4) <= '"&tyear2&"') GROUP BY userlist.emp_id, userlist.user_name, userlist.user_dept order by userlist.user_dept, userlist.emp_id  "
	Amount_MonthTotal = 0
	SellProfit_MonthTotal = 0
	SalesProfit_MonthTotal = 0

	Set rs6 = db.Execute(sqlPerson)
	
		Do until rs6.eof
			KRA = rs6("emp_id") '���
			Name = Trim(rs6("user_name"))'�̸�
			Team = Trim(rs6("user_dept"))'��
			amount_money_SUM_personal = 0
			Product_buy_price_Personal = 0
			exe_Rs_pay_amount_SUMPersonal = 0
			If Name <> "������" And Name <> "������" Then
			'�� ���� ���
			sqlPersonal = "SELECT SUM(amount_money) as sum, user_name from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& monthnow &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& monthnow &"-"& matchday2 &"' and user_name ='"&NAME&"' group by user_name"
			Set rs7 = db.Execute(sqlPersonal)
			If Not rs7.eof Then
			amount_money_SUM_personal = rs7("sum")
			End If

			sqlMinus = "SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& monthnow &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& monthnow &"-"& matchday2 &"' and user_name ='"&NAME&"' "
			Set rs8 = db.Execute(sqlMinus)
			Do until rs8.eof
			id = rs8("id")

			' ��ǰ����,���Ⱥ��
			'----------------------------------------------------------------------------
					sql = "select amount, account, account_des from ������� where id = '" &id& "' order by id"
					Set pro_Rs = db.Execute(sql)	
					
					' ��ǰ����, ��Ÿ���, ���Ⱥ�� ���
					Do until pro_Rs.eof
						amount=pro_Rs("amount") '�ݾ�
						account=Trim(pro_Rs("account"))'����
						account_des=Trim(pro_Rs("account_des"))'����

						IF account = "��ǰ����" Then
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						ElseIf account = "���Ⱥ��" Then
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						ELSEIF account_des = "��Ÿ" THEN 
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						End If
					pro_Rs.MoveNext
					Loop
			'���ְ��߿��� ���
					devmoneysql = "select tamount, services from ������ް�ȹ where id = '" &id& "' "
					Set pay_Rs = db.Execute(devmoneysql)	

					If Not pay_Rs.eof Then
						amount=pay_Rs("tamount") '�ݾ�
						services=Trim(pay_Rs("services")) '����

						IF services = "1" THEN 
						sell_pay_amount_SUM_Personal = sell_pay_amount_SUM_Personal + amount
						ELSEIF services = "2" THEN 
						sell_pay_amount_SUM_Personal = sell_pay_amount_SUM_Personal + amount
						END If
					End If
			' ���࿹��
			'----------------------------------------------------------------------------
					sql = "select sum(amount) as amount from ���࿹�� where id = '" &id& "'  group by id "
					Set exe_Rs = db.Execute(sql)	

					If Not exe_Rs.eof Then
					amount=exe_Rs("amount") '�ݾ�
					exe_Rs_pay_amount_SUM_Personal = exe_Rs_pay_amount_SUM_Personal + amount
					End If
					
					sql = "select sum(amount) as amount from �Ҹ𼺰�� where id = '" &id& "'   group by id "
					Set exe_Rs_1 = db.Execute(sql)	

					If Not exe_Rs_1.eof Then
					amount=exe_Rs_1("amount") '�ݾ�
					exe_Rs_1_pay_amount_SUM_Personal = exe_Rs_1_pay_amount_SUM_Personal + amount
					End If

					exe_Rs_pay_amount_SUMPersonal = exe_Rs_pay_amount_SUM_Personal + exe_Rs_1_pay_amount_SUM_Personal

				Product_buy_price_Personal = product_amount_SUM_Personal + sell_pay_amount_SUM_Personal
				Rs8.Movenext
				Loop	
%>
		<%If Dept <> Team Then%>
			<%If total = 1 Then%>
				<%Do Until TeamPersonCount = MaxPersonNum%>
					<tr> 
						<td class="a_right" style=" text-align: center">TBH<%=TeamPersonCount%></td>
						<td class="a_right"></td>
						<td class="a_right"></td>
						<td class="a_right"></td>
					</tr>
				<% TeamPersonCount = TeamPersonCount + 1
				Loop
				%>
			<tr>
				<th><font color="#0000FF"><b>�Ұ�</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(Product_buy_price_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(exe_Rs_pay_amount_SUMTeam_Total,0)%>��</b></th>
			</tr>
			<%End If%>
		<%
		Amount_MonthTotal = Amount_MonthTotal + amount_money_SUM_Team_Total
		SellProfit_MonthTotal = SellProfit_MonthTotal + Product_buy_price_Team_Total
		SalesProfit_MonthTotal = SalesProfit_MonthTotal + exe_Rs_pay_amount_SUMTeam_Total
		amount_money_SUM_Team_Total = 0
		Product_buy_price_Team_Total = 0
		exe_Rs_pay_amount_SUMTeam_Total = 0
		TeamPersonCount = 0
		%>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" class="tb_basic" style="width:400px; ">
			<tr> 
				<%If total = 0 Then%><th rowspan="<%=MaxPersonNum+3%>"><%=monthnow%>��</th><%End If%>
				<th colspan="4"><%=Team%></th>
			</tr>
			<tr> 
				<th>����</th>
				<th>����</th>
				<th>��������</th>
				<th>��������</th>
			</tr>
			<tr> 

				<td class="a_right" style=" text-align: center"><%=Name%></td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal,0)%>��</td>
			</tr>
		<%Else%>
			<tr>

				<td class="a_right" style=" text-align: center"><%=Name%></td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal,0)%>��</td>
			</tr>
		<%
		End If
		total = 1
		%>
		
<%
		product_amount_SUM_Personal = 0 
		sell_pay_amount_SUM_Personal = 0
		exe_Rs_pay_amount_SUM_Personal = 0
		exe_Rs_1_pay_amount_SUM_Personal = 0
		TeamPersonCount = TeamPersonCount + 1
		Dept = Team
		If Dept = Team Then
		amount_money_SUM_Team_Total = amount_money_SUM_Team_Total + amount_money_SUM_personal
		Product_buy_price_Team_Total = Product_buy_price_Team_Total + amount_money_SUM_personal - Product_buy_price_Personal
		exe_Rs_pay_amount_SUMTeam_Total = exe_Rs_pay_amount_SUMTeam_Total + amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal
		End If
		End If
		rs6.MoveNext
		Loop
%>
		<%Do Until TeamPersonCount = MaxPersonNum%>
			<tr> 
				<td class="a_right" style=" text-align: center">TBH<%=TeamPersonCount%></td>
				<td class="a_right"></td>
				<td class="a_right"></td>
				<td class="a_right"></td>
			</tr>
		<% TeamPersonCount = TeamPersonCount + 1
		Loop
		%>
			<tr>
				<th><font color="#0000FF"><b>�Ұ�</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(Product_buy_price_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(exe_Rs_pay_amount_SUMTeam_Total,0)%>��</b></th>
			</tr>
		
</div>

<%
'------------------�Ӵ�/�Ǽ����� ����----------------------------
	sqlPerson = "SELECT userlist.emp_id, userlist.user_name, userlist.user_dept FROM userlist INNER JOIN consultation ON userlist.emp_id = consultation.emp_id WHERE  (consultation.approval_flag = '1' or consultation.approval_flag = '3') and (left(consultation.contract_date_tax,4) >= '"&tyear1&"' AND left(consultation.contract_date_tax,4) <= '"&tyear2&"' and userlist.user_name = '������' or userlist.user_name = '������') GROUP BY userlist.emp_id, userlist.user_name, userlist.user_dept order by userlist.user_dept, userlist.emp_id  "
	
	Set rs6 = db.Execute(sqlPerson)
	
		Do until rs6.eof
			KRA = rs6("emp_id") '���
			Name = Trim(rs6("user_name"))'�̸�
			Team = Trim(rs6("user_dept"))'��


			'�� ���� ���
			sqlPersonal = "SELECT SUM(amount_money) as sum, user_name from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& monthnow &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& monthnow &"-"& matchday2 &"' and user_name ='"&NAME&"' group by user_name"
			Set rs7 = db.Execute(sqlPersonal)
			If Not rs7.eof Then
			amount_money_SUM_personal = rs7("sum")
			End If

			sqlMinus = "SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& monthnow &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& monthnow &"-"& matchday2 &"' and user_name ='"&NAME&"' "
			Set rs8 = db.Execute(sqlMinus)
			Do until rs8.eof
			id = rs8("id")

			' ��ǰ����,���Ⱥ��
			'----------------------------------------------------------------------------
					sql = "select amount, account, account_des from ������� where id = '" &id& "' order by id"
					Set pro_Rs = db.Execute(sql)	
					
					' ��ǰ����, ��Ÿ���, ���Ⱥ�� ���
					Do until pro_Rs.eof
						amount=pro_Rs("amount") '�ݾ�
						account=Trim(pro_Rs("account"))'����
						account_des=Trim(pro_Rs("account_des"))'����

						IF account = "��ǰ����" Then
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						ElseIf account = "���Ⱥ��" Then
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						ELSEIF account_des = "��Ÿ" THEN 
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						End If
					pro_Rs.MoveNext
					Loop
			'���ְ��߿��� ���
					devmoneysql = "select tamount, services from ������ް�ȹ where id = '" &id& "' "
					Set pay_Rs = db.Execute(devmoneysql)	

					If Not pay_Rs.eof Then
						amount=pay_Rs("tamount") '�ݾ�
						services=Trim(pay_Rs("services")) '����

						IF services = "1" THEN 
						sell_pay_amount_SUM_Personal = sell_pay_amount_SUM_Personal + amount
						ELSEIF services = "2" THEN 
						sell_pay_amount_SUM_Personal = sell_pay_amount_SUM_Personal + amount
						END If
					End If
			' ���࿹��
			'----------------------------------------------------------------------------
					sql = "select sum(amount) as amount from ���࿹�� where id = '" &id& "'  group by id "
					Set exe_Rs = db.Execute(sql)	

					If Not exe_Rs.eof Then
					amount=exe_Rs("amount") '�ݾ�
					exe_Rs_pay_amount_SUM_Personal = exe_Rs_pay_amount_SUM_Personal + amount
					End If
					
					sql = "select sum(amount) as amount from �Ҹ𼺰�� where id = '" &id& "'   group by id "
					Set exe_Rs_1 = db.Execute(sql)	

					If Not exe_Rs_1.eof Then
					amount=exe_Rs_1("amount") '�ݾ�
					exe_Rs_1_pay_amount_SUM_Personal = exe_Rs_1_pay_amount_SUM_Personal + amount
					End If

					exe_Rs_pay_amount_SUMPersonal = exe_Rs_pay_amount_SUM_Personal + exe_Rs_1_pay_amount_SUM_Personal

				Product_buy_price_Personal = product_amount_SUM_Personal + sell_pay_amount_SUM_Personal
				Rs8.Movenext
				Loop	
%>
		<%If Dept <> Team  Then%>
			<%If total = 2 Then%>
				<%Do Until TeamPersonCount = MaxPersonNum%>
				<tr> 
					<td class="a_right" style=" text-align: center">TBH<%=TeamPersonCount%></td>
					<td class="a_right"></td>
					<td class="a_right"></td>
					<td class="a_right"></td>
				</tr>
				<% TeamPersonCount = TeamPersonCount + 1
				Loop
				%>
			<tr>
				<th><font color="#0000FF"><b>�Ұ�<%=TeamPersonCount%></b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(Product_buy_price_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(exe_Rs_pay_amount_SUMTeam_Total,0)%>��</b></th>
			</tr>
			<%End If%>
		<%
		Amount_MonthTotal = Amount_MonthTotal + amount_money_SUM_Team_Total
		SellProfit_MonthTotal = SellProfit_MonthTotal + Product_buy_price_Team_Total
		SalesProfit_MonthTotal = SalesProfit_MonthTotal + exe_Rs_pay_amount_SUMTeam_Total
		amount_money_SUM_Team_Total = 0
		Product_buy_price_Team_Total = 0
		exe_Rs_pay_amount_SUMTeam_Total = 0
		TeamPersonCount = 0
		%>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" class="tb_basic" style="width:400px; ">
			<tr> 
				<!---<th rowspan="<%=MaxPersonNum+3%>"><%=monthnow%>��</th>--->
				<th colspan="4">
					<% If Name = "������" Then%>
					�Ӵ�(������)
					<% ElseIf Name = "������" Then%>
					�Ǽ����������
					<%End If%>
				</th>
			</tr>
			<tr> 
				<th>����</th>
				<th>����</th>
				<th>��������</th>
				<th>��������</th>
			</tr>
			<tr> 
				<td class="a_right" style=" text-align: center"><%=Name%></td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal,0)%>��</td>
			</tr>
		<%Else%>
			<tr>
				<td class="a_right" style=" text-align: center"><%=Name%></td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal,0)%>��</td>
			</tr>
		<%
		End If
		total = 2
		%>
<%
		
		product_amount_SUM_Personal = 0 
		sell_pay_amount_SUM_Personal = 0
		exe_Rs_pay_amount_SUM_Personal = 0
		exe_Rs_1_pay_amount_SUM_Personal = 0
		TeamPersonCount = TeamPersonCount + 1
		Dept = Team
		If Dept = Team Then
		amount_money_SUM_Team_Total = amount_money_SUM_Team_Total + amount_money_SUM_personal
		Product_buy_price_Team_Total = Product_buy_price_Team_Total + amount_money_SUM_personal - Product_buy_price_Personal
		exe_Rs_pay_amount_SUMTeam_Total = exe_Rs_pay_amount_SUMTeam_Total + amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal
		End If

		rs6.MoveNext
		Loop
		'------------------�Ӵ�/�Ǽ����� ���� ��----------------------------
%>
		<!------------------�Ӵ�/�Ǽ����� ������ ���̺� ������------------------->
		<%Do Until TeamPersonCount = MaxPersonNum%>
			<tr> 
				<td class="a_right" style=" text-align: center">TBH<%=TeamPersonCount%></td>
				<td class="a_right"></td>
				<td class="a_right"></td>
				<td class="a_right"></td>
			</tr>
		<% TeamPersonCount = TeamPersonCount + 1
		Loop
		%>
			<tr>
				<th><font color="#0000FF"><b>�Ұ�<%=TeamPersonCount%></b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(Product_buy_price_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(exe_Rs_pay_amount_SUMTeam_Total,0)%>��</b></th>
			</tr>
			<%
				Amount_MonthTotal = Amount_MonthTotal + amount_money_SUM_Team_Total
				SellProfit_MonthTotal = SellProfit_MonthTotal + Product_buy_price_Team_Total
				SalesProfit_MonthTotal = SalesProfit_MonthTotal + exe_Rs_pay_amount_SUMTeam_Total
			%>
			</table>
			<!---�հ� �� �˻�--->
			<table border="0" cellpadding="0" cellspacing="0" class="tb_basic" style="width:400px; ">
			<tr> 

				<th colspan="3">�հ� �� �˻�</th>
			</tr>
			<tr> 
				<th>����</th>
				<th>��������</th>
				<th>��������</th>
			</tr>
			<%
			Cnt = 0
			Do Until Cnt = MaxPersonNum%>
			<tr> 
				<td style="text-align: center"><font color="#0000FF"><b>N/A</b></td>
				<td style="text-align: center"><font color="#0000FF"><b>N/A</b></td>
				<td style="text-align: center"><font color="#0000FF"><b>N/A</b></td>
			</tr>
			<% Cnt = Cnt + 1
			Loop
			%>
			<tr> 
				<th><font color="#0000FF"><b><%=FormatNumber(Amount_MonthTotal,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(SellProfit_MonthTotal,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(SalesProfit_MonthTotal,0)%>��</b></th>
			</tr>
			</table>
			<div style="float:left;clear:both;width:100%;height:2px;background:black;"></div>
<%
			If currentmonth = month1 Then
				monthnow = month2
				currentmonth = month2
			ElseIf currentmonth = month2 Then
				monthnow = month3
			End If
			months = months + 1
			total = 0
			loop
			%>

<!--- --------------------���� �б� �հ� ---------------------->

<%	
	sqlPersonSum = "SELECT Top 1 count(a.user_dept) as max FROM (SELECT userlist.emp_id, userlist.user_name, userlist.user_dept FROM userlist INNER JOIN consultation ON userlist.emp_id = consultation.emp_id WHERE  (consultation.approval_flag = '1' or consultation.approval_flag = '3') and (left(consultation.contract_date_tax,4) >= '"&tyear1&"' AND left(consultation.contract_date_tax,4) <= '"&tyear2&"') GROUP BY userlist.emp_id, userlist.user_name, userlist.user_dept ) a GROUP BY a.user_dept order by max desc "
		Set personNum = db.Execute(sqlPersonSum)
			If Not personNum.eof Then
				MaxPersonNum = personNum("max")
			End If

	'���κ� ����


	sqlPerson = "SELECT userlist.emp_id, userlist.user_name, userlist.user_dept FROM userlist INNER JOIN consultation ON userlist.emp_id = consultation.emp_id WHERE  (consultation.approval_flag = '1' or consultation.approval_flag = '3') and (left(consultation.contract_date_tax,4) >= '"&tyear1&"' AND left(consultation.contract_date_tax,4) <= '"&tyear2&"') GROUP BY userlist.emp_id, userlist.user_name, userlist.user_dept order by userlist.user_dept, userlist.emp_id  "
	Amount_MonthTotal = 0
	SellProfit_MonthTotal = 0
	SalesProfit_MonthTotal = 0

	Set rs6 = db.Execute(sqlPerson)
	
		Do until rs6.eof
			KRA = rs6("emp_id") '���
			Name = Trim(rs6("user_name"))'�̸�
			Team = Trim(rs6("user_dept"))'��
			amount_money_SUM_personal = 0
			Product_buy_price_Personal = 0
			exe_Rs_pay_amount_SUMPersonal = 0
			If Name <> "������" And Name <> "������" Then
			'�� ���� ���
			sqlPersonal = "SELECT SUM(amount_money) as sum, user_name from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month1 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month3 &"-"& matchday2 &"' and user_name ='"&NAME&"' group by user_name"
			Set rs7 = db.Execute(sqlPersonal)
			If Not rs7.eof Then
			amount_money_SUM_personal = rs7("sum")
			End If

			sqlMinus = "SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month1 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month3 &"-"& matchday2 &"' and user_name ='"&NAME&"' "
			Set rs8 = db.Execute(sqlMinus)
			Do until rs8.eof
			id = rs8("id")

			' ��ǰ����,���Ⱥ��
			'----------------------------------------------------------------------------
					sql = "select amount, account, account_des from ������� where id = '" &id& "' order by id"
					Set pro_Rs = db.Execute(sql)	
					
					' ��ǰ����, ��Ÿ���, ���Ⱥ�� ���
					Do until pro_Rs.eof
						amount=pro_Rs("amount") '�ݾ�
						account=Trim(pro_Rs("account"))'����
						account_des=Trim(pro_Rs("account_des"))'����

						IF account = "��ǰ����" Then
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						ElseIf account = "���Ⱥ��" Then
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						ELSEIF account_des = "��Ÿ" THEN 
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						End If
					pro_Rs.MoveNext
					Loop
			'���ְ��߿��� ���
					devmoneysql = "select tamount, services from ������ް�ȹ where id = '" &id& "' "
					Set pay_Rs = db.Execute(devmoneysql)	

					If Not pay_Rs.eof Then
						amount=pay_Rs("tamount") '�ݾ�
						services=Trim(pay_Rs("services")) '����

						IF services = "1" THEN 
						sell_pay_amount_SUM_Personal = sell_pay_amount_SUM_Personal + amount
						ELSEIF services = "2" THEN 
						sell_pay_amount_SUM_Personal = sell_pay_amount_SUM_Personal + amount
						END If
					End If
			' ���࿹��
			'----------------------------------------------------------------------------
					sql = "select sum(amount) as amount from ���࿹�� where id = '" &id& "'  group by id "
					Set exe_Rs = db.Execute(sql)	

					If Not exe_Rs.eof Then
					amount=exe_Rs("amount") '�ݾ�
					exe_Rs_pay_amount_SUM_Personal = exe_Rs_pay_amount_SUM_Personal + amount
					End If
					
					sql = "select sum(amount) as amount from �Ҹ𼺰�� where id = '" &id& "'   group by id "
					Set exe_Rs_1 = db.Execute(sql)	

					If Not exe_Rs_1.eof Then
					amount=exe_Rs_1("amount") '�ݾ�
					exe_Rs_1_pay_amount_SUM_Personal = exe_Rs_1_pay_amount_SUM_Personal + amount
					End If

					exe_Rs_pay_amount_SUMPersonal = exe_Rs_pay_amount_SUM_Personal + exe_Rs_1_pay_amount_SUM_Personal

				Product_buy_price_Personal = product_amount_SUM_Personal + sell_pay_amount_SUM_Personal
				Rs8.Movenext
				Loop	
%>
		<%If Dept <> Team Then%>
			<%If total = 1 Then%>
				<%Do Until TeamPersonCount = MaxPersonNum%>
					<tr> 
						<td class="a_right" style=" text-align: center">TBH<%=TeamPersonCount%></td>
						<td class="a_right"></td>
						<td class="a_right"></td>
						<td class="a_right"></td>
					</tr>
				<% TeamPersonCount = TeamPersonCount + 1
				Loop
				%>
			<tr>
				<th><font color="#0000FF"><b>�Ұ�</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(Product_buy_price_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(exe_Rs_pay_amount_SUMTeam_Total,0)%>��</b></th>
			</tr>
			<%End If%>
		<%
		Amount_MonthTotal = Amount_MonthTotal + amount_money_SUM_Team_Total
		SellProfit_MonthTotal = SellProfit_MonthTotal + Product_buy_price_Team_Total
		SalesProfit_MonthTotal = SalesProfit_MonthTotal + exe_Rs_pay_amount_SUMTeam_Total
		amount_money_SUM_Team_Total = 0
		Product_buy_price_Team_Total = 0
		exe_Rs_pay_amount_SUMTeam_Total = 0
		TeamPersonCount = 0
		%>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" class="tb_basic" style="width:400px; ">
			<tr> 
				<%If total = 0 Then%><th rowspan="<%=MaxPersonNum+3%>">&nbsp;Q<%=matchQuarter%>&nbsp;</th><%End If%>
				<th colspan="4"><%=Team%></th>
			</tr>
			<tr> 
				<th>����</th>
				<th>����</th>
				<th>��������</th>
				<th>��������</th>
			</tr>
			<tr> 

				<td class="a_right" style=" text-align: center"><%=Name%></td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal,0)%>��</td>
			</tr>
		<%Else%>
			<tr>

				<td class="a_right" style=" text-align: center"><%=Name%></td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal,0)%>��</td>
			</tr>
		<%
		End If
		total = 1
		%>
		
<%
		product_amount_SUM_Personal = 0 
		sell_pay_amount_SUM_Personal = 0
		exe_Rs_pay_amount_SUM_Personal = 0
		exe_Rs_1_pay_amount_SUM_Personal = 0
		TeamPersonCount = TeamPersonCount + 1
		Dept = Team
		If Dept = Team Then
		amount_money_SUM_Team_Total = amount_money_SUM_Team_Total + amount_money_SUM_personal
		Product_buy_price_Team_Total = Product_buy_price_Team_Total + amount_money_SUM_personal - Product_buy_price_Personal
		exe_Rs_pay_amount_SUMTeam_Total = exe_Rs_pay_amount_SUMTeam_Total + amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal
		End If
		End If
		rs6.MoveNext
		Loop
%>
		<%Do Until TeamPersonCount = MaxPersonNum%>
			<tr> 
				<td class="a_right" style=" text-align: center">TBH<%=TeamPersonCount%></td>
				<td class="a_right"></td>
				<td class="a_right"></td>
				<td class="a_right"></td>
			</tr>
		<% TeamPersonCount = TeamPersonCount + 1
		Loop
		%>
			<tr>
				<th><font color="#0000FF"><b>�Ұ�</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(Product_buy_price_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(exe_Rs_pay_amount_SUMTeam_Total,0)%>��</b></th>
			</tr>
		
</div>

<%
'------------------�Ӵ�/�Ǽ����� ����----------------------------
	sqlPerson = "SELECT userlist.emp_id, userlist.user_name, userlist.user_dept FROM userlist INNER JOIN consultation ON userlist.emp_id = consultation.emp_id WHERE  (consultation.approval_flag = '1' or consultation.approval_flag = '3') and (left(consultation.contract_date_tax,4) >= '"&tyear1&"' AND left(consultation.contract_date_tax,4) <= '"&tyear2&"' and userlist.user_name = '������' or userlist.user_name = '������') GROUP BY userlist.emp_id, userlist.user_name, userlist.user_dept order by userlist.user_dept, userlist.emp_id  "
	
	Set rs6 = db.Execute(sqlPerson)
	
		Do until rs6.eof
			KRA = rs6("emp_id") '���
			Name = Trim(rs6("user_name"))'�̸�
			Team = Trim(rs6("user_dept"))'��


			'�� ���� ���
			sqlPersonal = "SELECT SUM(amount_money) as sum, user_name from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month1 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month3 &"-"& matchday2 &"' and user_name ='"&NAME&"' group by user_name"
			Set rs7 = db.Execute(sqlPersonal)
			If Not rs7.eof Then
			amount_money_SUM_personal = rs7("sum")
			End If

			sqlMinus = "SELECT  client, amount_money, id, w_date, riskrate, project_code1, project_code2, project_name,customer_name,user_name,contract_date ,contract_date_tax from ConsultationTeam_List s1 where s1.versionid =( select Max(versionid) from ConsultationTeam_List s2 where s2.project_code1=s1.project_code1 and s2.project_code2=s1.project_code2 and s2.doc_kind<>'15' and s2.doc_kind<>'16' and (s2.approval_flag='1' or  s2.approval_flag='3')) and contract_date_tax>='"&tyear1&"-"& month1 &"-"& matchday1 &"' and contract_date_tax<='"&tyear2&"-"& month3 &"-"& matchday2 &"' and user_name ='"&NAME&"' "
			Set rs8 = db.Execute(sqlMinus)
			Do until rs8.eof
			id = rs8("id")

			' ��ǰ����,���Ⱥ��
			'----------------------------------------------------------------------------
					sql = "select amount, account, account_des from ������� where id = '" &id& "' order by id"
					Set pro_Rs = db.Execute(sql)	
					
					' ��ǰ����, ��Ÿ���, ���Ⱥ�� ���
					Do until pro_Rs.eof
						amount=pro_Rs("amount") '�ݾ�
						account=Trim(pro_Rs("account"))'����
						account_des=Trim(pro_Rs("account_des"))'����

						IF account = "��ǰ����" Then
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						ElseIf account = "���Ⱥ��" Then
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						ELSEIF account_des = "��Ÿ" THEN 
						product_amount_SUM_Personal = product_amount_SUM_Personal + amount
						End If
					pro_Rs.MoveNext
					Loop
			'���ְ��߿��� ���
					devmoneysql = "select tamount, services from ������ް�ȹ where id = '" &id& "' "
					Set pay_Rs = db.Execute(devmoneysql)	

					If Not pay_Rs.eof Then
						amount=pay_Rs("tamount") '�ݾ�
						services=Trim(pay_Rs("services")) '����

						IF services = "1" THEN 
						sell_pay_amount_SUM_Personal = sell_pay_amount_SUM_Personal + amount
						ELSEIF services = "2" THEN 
						sell_pay_amount_SUM_Personal = sell_pay_amount_SUM_Personal + amount
						END If
					End If
			' ���࿹��
			'----------------------------------------------------------------------------
					sql = "select sum(amount) as amount from ���࿹�� where id = '" &id& "'  group by id "
					Set exe_Rs = db.Execute(sql)	

					If Not exe_Rs.eof Then
					amount=exe_Rs("amount") '�ݾ�
					exe_Rs_pay_amount_SUM_Personal = exe_Rs_pay_amount_SUM_Personal + amount
					End If
					
					sql = "select sum(amount) as amount from �Ҹ𼺰�� where id = '" &id& "'   group by id "
					Set exe_Rs_1 = db.Execute(sql)	

					If Not exe_Rs_1.eof Then
					amount=exe_Rs_1("amount") '�ݾ�
					exe_Rs_1_pay_amount_SUM_Personal = exe_Rs_1_pay_amount_SUM_Personal + amount
					End If

					exe_Rs_pay_amount_SUMPersonal = exe_Rs_pay_amount_SUM_Personal + exe_Rs_1_pay_amount_SUM_Personal

				Product_buy_price_Personal = product_amount_SUM_Personal + sell_pay_amount_SUM_Personal
				Rs8.Movenext
				Loop	
%>
		<%If Dept <> Team  Then%>
			<%If total = 2 Then%>
				<%Do Until TeamPersonCount = MaxPersonNum%>
				<tr> 
					<td class="a_right" style=" text-align: center">TBH<%=TeamPersonCount%></td>
					<td class="a_right"></td>
					<td class="a_right"></td>
					<td class="a_right"></td>
				</tr>
				<% TeamPersonCount = TeamPersonCount + 1
				Loop
				%>
			<tr>
				<th><font color="#0000FF"><b>�Ұ�<%=TeamPersonCount%></b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(Product_buy_price_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(exe_Rs_pay_amount_SUMTeam_Total,0)%>��</b></th>
			</tr>
			<%End If%>
		<%
		Amount_MonthTotal = Amount_MonthTotal + amount_money_SUM_Team_Total
		SellProfit_MonthTotal = SellProfit_MonthTotal + Product_buy_price_Team_Total
		SalesProfit_MonthTotal = SalesProfit_MonthTotal + exe_Rs_pay_amount_SUMTeam_Total
		amount_money_SUM_Team_Total = 0
		Product_buy_price_Team_Total = 0
		exe_Rs_pay_amount_SUMTeam_Total = 0
		TeamPersonCount = 0
		%>
		</table>
		<table border="0" cellpadding="0" cellspacing="0" class="tb_basic" style="width:400px; ">
			<tr> 
				<!---<th rowspan="<%=MaxPersonNum+3%>">&nbsp;Q<%=matchQuarter%>&nbsp;</th>--->
				<th colspan="4">
					<% If Name = "������" Then%>
					�Ӵ�(������)
					<% ElseIf Name = "������" Then%>
					�Ǽ����������
					<%End If%>
				</th>
			</tr>
			<tr> 
				<th>����</th>
				<th>����</th>
				<th>��������</th>
				<th>��������</th>
			</tr>
			<tr> 
				<td class="a_right" style=" text-align: center"><%=Name%></td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal,0)%>��</td>
			</tr>
		<%Else%>
			<tr>
				<td class="a_right" style=" text-align: center"><%=Name%></td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal,0)%>��</td>
				<td class="a_right"><%=FormatNumber(amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal,0)%>��</td>
			</tr>
		<%
		End If
		total = 2
		%>
<%
		
		product_amount_SUM_Personal = 0 
		sell_pay_amount_SUM_Personal = 0
		exe_Rs_pay_amount_SUM_Personal = 0
		exe_Rs_1_pay_amount_SUM_Personal = 0
		TeamPersonCount = TeamPersonCount + 1
		Dept = Team
		If Dept = Team Then
		amount_money_SUM_Team_Total = amount_money_SUM_Team_Total + amount_money_SUM_personal
		Product_buy_price_Team_Total = Product_buy_price_Team_Total + amount_money_SUM_personal - Product_buy_price_Personal
		exe_Rs_pay_amount_SUMTeam_Total = exe_Rs_pay_amount_SUMTeam_Total + amount_money_SUM_personal - Product_buy_price_Personal - exe_Rs_pay_amount_SUMPersonal
		End If

		rs6.MoveNext
		Loop
		'------------------�Ӵ�/�Ǽ����� ���� ��----------------------------
%>
		<!------------------�Ӵ�/�Ǽ����� ������ ���̺� ������------------------->
		<%Do Until TeamPersonCount = MaxPersonNum%>
			<tr> 
				<td class="a_right" style=" text-align: center">TBH<%=TeamPersonCount%></td>
				<td class="a_right"></td>
				<td class="a_right"></td>
				<td class="a_right"></td>
			</tr>
		<% TeamPersonCount = TeamPersonCount + 1
		Loop
		%>
			<tr>
				<th><font color="#0000FF"><b>�Ұ�<%=TeamPersonCount%></b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(Product_buy_price_Team_Total,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(exe_Rs_pay_amount_SUMTeam_Total,0)%>��</b></th>
			</tr>
			<%
				Amount_MonthTotal = Amount_MonthTotal + amount_money_SUM_Team_Total
				SellProfit_MonthTotal = SellProfit_MonthTotal + Product_buy_price_Team_Total
				SalesProfit_MonthTotal = SalesProfit_MonthTotal + exe_Rs_pay_amount_SUMTeam_Total
			%>
			</table>
			<!---�հ� �� �˻�--->
			<table border="0" cellpadding="0" cellspacing="0" class="tb_basic" style="width:400px; ">
			<tr> 

				<th colspan="3">�հ� �� �˻�</th>
			</tr>
			<tr> 
				<th>����</th>
				<th>��������</th>
				<th>��������</th>
			</tr>
			<%
			Cnt = 0
			Do Until Cnt = MaxPersonNum%>
			<tr> 
				<td style="text-align: center"><font color="#0000FF"><b>N/A</b></td>
				<td style="text-align: center"><font color="#0000FF"><b>N/A</b></td>
				<td style="text-align: center"><font color="#0000FF"><b>N/A</b></td>
			</tr>
			<% Cnt = Cnt + 1
			Loop
			%>
			<tr> 
				<th><font color="#0000FF"><b><%=FormatNumber(Amount_MonthTotal,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(SellProfit_MonthTotal,0)%>��</b></th>
				<th><font color="#0000FF"><b><%=FormatNumber(SalesProfit_MonthTotal,0)%>��</b></th>
			</tr>
			</table>
<%

			total = 0

			%>