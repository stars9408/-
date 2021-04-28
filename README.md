# 손익계산서 페이지
그룹웨어 내에서 DB를 활용하여 연별/분기별/월별 매출, 매입, 이익을 계산하는 페이지입니다.

2021년 4월 현재 직장에서 내부 개발로 맡아 진행했던 경험입니다.

https://github.com/stars9408/GW_ProfitCalculator

웹 페이지 주소: 내부 그룹웨어

---

# 요구사항:

1. 선택한 분기에 알맞는 전사/비교/팀별 손익계산서를 만들어야 한다.
    1. 테이블 1: **전사** 분기 내 월별/분기종합 **매출 및 영업익**
    2. 테이블 2: 사내 분기 **비교 매출 및 영업익** 
    3. 테이블 3: **팀별/개인별/월별** **매출, 매출이익, 영업이익**
    4. 테이블 4: **팀별/분기별 매출, 매출이익, 영업이익**
2. 아래 사항들을 충족해야 합니다.
    1. 특정 인물들에 대하여, 팀별 자료에서 제외하여 따로 새로운 테이블을 생성하여 자료를 출력한다.
    2. 월별자료 끝에는 모든팀의 월별 합을 계산하여 출력한다.
    3. 테이블의 크기가 일정하도록 가장 많은 팀원을 가진 팀의 인원 수 만큼 TBH 인원을 각 팀 테이블에 채워준다.

# 활용 Skills:

1. ASP-HTML
    1. FORM, VBSCRIPT 등
2. JAVASCRIPT
3. MSSQL DB (Microsoft DB Server Management Service)

# 프로세스:

1. **사전 기능 추가 사항**
    1. 회사 그룹웨어 사이드 바 內 링크 추가
        1. 링크 내 손익계산서 페이지 개발
    2. 상단 연도 및 분기 선택 옵션 추가 & 엑셀 다운로드 버튼 추가
2. **테이블 1: 전사 분기 내 월별/분기종합 매출 및 영업익**
    1. 선택된 분기의 전사 월별 사업 실적 출력
        1. Quarter에 따른 월별 매출 및 영업익 출력
            1. 해당 DB의 테이블 및 뷰를 활용하여, 매출, 매입가, 프로젝트 비용, 인건비, 기타 비용 등을 계산하여 영업익을 계산
            2. 3개의 월별 자료 및 해당 Q의 합계 출력
3. **테이블 2: 사내 분기 비교 매출 및 영업익** 
    1. 당사 내 비교자료
    2. 해당 분기(Q)에 따른 1사와 2사의 매출과 영업이익 및 합계 출력
4. **테이블 3: 팀별/개인별/월별 매출, 매출이익, 영업이익**
    1. 선택된 분기(Q)에 알맞는 월별 자료를 팀별로 나누어 각각 개개인의 매출, 매출이익, 영업이익을 계산하여 출력
    2. *각 팀마다 인원 수가 상이하므로, 가장 긴 테이블(가장 인원이 많은 팀)의 세로 길이에 맞추어 모든 팀에서 같은 크기의 테이블을 가지도록 한다.
        1. TBH1...TBH4 등의 이름으로 비어있는 ROW를 생성하여 삽입.
    3. 특정 인원의 매출 및 이익 자료를 해당 팀에서가 아닌 개인의 테이블로 새로 지정하여 생성해준다(총 2인).
    4. 맨 끝에는 새로운 테이블을 생성하여 해당 월의 모든 팀의 매출, 매출이익, 영업이익 합계 자료를 출력한다.
5. **테이블 4: 팀별/분기별 매출, 매출이익, 영업이익**
    1. 맨 밑에는 새로운 테이블들을 생성하여 각 팀별 및 특정 개인별 분기 매출, 매출이익, 영업이익 합계를 출력한다.
    2. 맨 끝에는 새로운 테이블을 생성하여 해당 분기의 모든 팀의 매출, 매출이익, 영업이익 합계 자료를 출력한다.

## 그룹웨어 메인 사이드바 링크:

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/c40e3316-002b-4aaa-adb9-b8941f19214a/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/c40e3316-002b-4aaa-adb9-b8941f19214a/Untitled.png)

## **초안 엑셀 파일 모습:**

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/6bf4da0e-3906-42aa-9d2a-82f0380014a3/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/6bf4da0e-3906-42aa-9d2a-82f0380014a3/Untitled.png)

- **분기 선택 기능:**

```bash
**먼저 선택된 분기에 따라 월/일을 설정해주는 코드입니다.**

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
```

- **아래와 같이 SQL문 등을 활용하여 매출, 매출이익, 영업이익을 계산하였습니다. (*다른 계산 코드는 포함하지 않았습니다.**

```jsx
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
	'달 계산
	Do until rs.eof   
			id = rs("id")
			user_name = rs("user_name")
			contract_date_tax = rs("contract_date_tax")
			amount_money = rs("amount_money")

			' 상품원가,제안비용
			'----------------------------------------------------------------------------
					sql = "select amount, account, account_des from 매출원가 where id = '" &id& "' "
					Set pro_Rs = db.Execute(sql)	
					
					' 상품원가, 기타비용, 제안비용 계산
					Do until pro_Rs.eof
						amount=pro_Rs("amount") '금액
						account=Trim(pro_Rs("account"))'종류
						account_des=Trim(pro_Rs("account_des"))'종류

						IF account = "상품원가" Then
						product_amount_SUM_1 = product_amount_SUM_1 + amount
						ElseIf account = "제안비용" Then
						product_amount_SUM_1 = product_amount_SUM_1 + amount
						ELSEIF account_des = "기타" THEN 
						product_amount_SUM_1 = product_amount_SUM_1 + amount
						End If
					pro_Rs.MoveNext
					Loop

					'외주개발원가 계산
					devmoneysql = "select tamount, services from 대금지급계획 where id = '" &id& "' "
					Set pay_Rs = db.Execute(devmoneysql)	

					If Not pay_Rs.eof Then
						amount=pay_Rs("tamount") '금액
						services=Trim(pay_Rs("services")) '종류

						IF services = "1" THEN 
						sell_pay_amount_SUM_1 = sell_pay_amount_SUM_1 + amount
						ELSEIF services = "2" THEN 
						sell_pay_amount_SUM_1 = sell_pay_amount_SUM_1 + amount
						END If
					End If
					
					' 실행예산
			'----------------------------------------------------------------------------
					sql = "select sum(amount) as amount from 실행예산 where id = '" &id& "'  group by id "
					Set exe_Rs = db.Execute(sql)	

					If Not exe_Rs.eof Then
					amount=exe_Rs("amount") '금액
					exe_Rs_pay_amount_SUM_1 = exe_Rs_pay_amount_SUM_1 + amount
					End If

					
					sql = "select sum(amount) as amount from 소모성경비 where id = '" &id& "'   group by id "
					Set exe_Rs_1 = db.Execute(sql)	

					If Not exe_Rs_1.eof Then
					amount=exe_Rs_1("amount") '금액
					exe_Rs_1_pay_amount_SUM_1 = exe_Rs_1_pay_amount_SUM_1 + amount
					End If

					exe_Rs_pay_amount_SUM = exe_Rs_pay_amount_SUM_1 + exe_Rs_1_pay_amount_SUM_1
			'달별 매출, 영업익
			amount_money_SUM_TOT = amount_money_SUM_TOT + amount_money
			amount_money_total = amount_money_SUM_TOT - product_amount_SUM_1 - sell_pay_amount_SUM_1 - exe_Rs_pay_amount_SUM

			Rs.Movenext
			Loop

			'쿼터 매출, 영업익 합계
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

			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_SUM_TOT,0)%>원</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(amount_money_total,0)%>원</b></td>	

			<%
			Loop	
			%>

			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(Total_amount_money_SUM,0)%>원</b></td>
			<td class="a_right" style="text-align:center;"><font color="#0000FF"><b><%=FormatNumber(Total_earn_money_Sum,0)%>원</b></td>
		</tr>
		</table>
```

## 손익계산서 페이지

### 상단 연도/분기 선택 및 엑셀 다운로드 버튼

- 선택된 분기 및 현재 날짜 표시

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/10ba3bd8-1eed-4f3f-bb7b-c8fe7e99f494/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/10ba3bd8-1eed-4f3f-bb7b-c8fe7e99f494/Untitled.png)

### 테이블 1:

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/4f3992a4-ab4f-4d11-9c8a-6f9e9a1f2d8e/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/4f3992a4-ab4f-4d11-9c8a-6f9e9a1f2d8e/Untitled.png)

### 테이블 2:

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/3bbf1ad9-f845-4c96-b247-49aa67c72085/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/3bbf1ad9-f845-4c96-b247-49aa67c72085/Untitled.png)

### 테이블 3:

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/3ce0b6e0-7a77-4eaf-9815-ea5ea4b5dbea/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/3ce0b6e0-7a77-4eaf-9815-ea5ea4b5dbea/Untitled.png)

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/3881e525-fb69-44ab-ad9e-69779c262d2f/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/3881e525-fb69-44ab-ad9e-69779c262d2f/Untitled.png)

### 테이블 4:

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/1fecff65-2dc4-4ea6-9c06-0690f28c9e17/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/1fecff65-2dc4-4ea6-9c06-0690f28c9e17/Untitled.png)

![https://s3-us-west-2.amazonaws.com/secure.notion-static.com/8a94ad78-440e-4d35-a00f-168fc70d6299/Untitled.png](https://s3-us-west-2.amazonaws.com/secure.notion-static.com/8a94ad78-440e-4d35-a00f-168fc70d6299/Untitled.png)

# 또 다른 습득

- **JOIN 및 FROM ( SELECT ...)의 더 다양한 방식으로 SQL문을 활용하여 DB를 사용하였습니다.**
- **더 적은 수의 SQL EXECUTION으로 더 많은 부분들을 계산하기 위해, IF, DO WHILE 을 여러방면으로 활용하여, 보다 빠른 알고리즘을 고안하였습니다.**

# 느낌점 및 후기:

- **그저 쉽게 개발하기 위해서 많은 SQL문을 EXECUTE하여 만들 수 있었지만, 더 나은 알고리즘과 더 빠른 기능을 위해 여러번 사고하고 고민하여 IF 문과 DO WHILE문을 사용하였습니다. SQL EXECUTION을 하나 하나 줄여 갈 때마다 성취감을 느낄 수 있었고, 페이지 로딩 또한 줄어드는 것을 볼 수 있었습니다.**
- **DB를 활용함에 있어서도, 더욱 다양한 표현법과 방식을 사용하여 여러개의 테이블을 하나의 SQL에서 운용할 수 있었습니다. 그렇기에 더욱 더 간결한 알고리즘을 만들어낼 수 있었습니다.**
