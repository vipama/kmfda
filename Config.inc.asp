<meta http-equiv="Content-Type" content="text/html; charset=windows-874">
<%@Language="VBScript.Encode" %>
<%
	'ส่วนติดต่อฐานข้อมูล
	'Don't Fix==================================
		Set Con = Server.CreateObject("ADODB.Connection")
		'Str_Connect = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source ="&Server.mappath("\030_KM\_db\KM.mdb")
		Str_Connect = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source =E:\_db\KM.mdb"
		'Str_Connect = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source ="&Server.MapPath("..\_db\KM.mdb")
		
    	'Str_Connect = "Provider=SQLOLEDB;Server=(local);integrated security=SSPI;uid=sa;pwd=coolooc;database=library"
		'Str_Connect = "driver={SQL Server};server=localhost;uid=sa;pwd=coolooc;database=library;" 
		Con.open Str_Connect
		
		
		Set ConQS = Server.CreateObject("ADODB.Connection")
		'Str_Connect = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source ="&Server.mappath("\030_KM\_db\KM.mdb")
		Str_ConnectQS = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source =E:\_db\QS_DB.mdb"
		'Str_ConnectQS = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source ="&Server.MapPath("..\_db\QS_DB.mdb")
		
    	'Str_Connect = "Provider=SQLOLEDB;Server=(local);integrated security=SSPI;uid=sa;pwd=coolooc;database=library"
		'Str_Connect = "driver={SQL Server};server=localhost;uid=sa;pwd=coolooc;database=library;" 
		ConQS.open Str_ConnectQS
		
		
		Set ConSql2000 = Server.CreateObject("ADODB.Connection")
		'Str_Connect2000 = "PROVIDER=SQLOLEDB; DATA SOURCE=(local); INITIAL CATALOG=weblibrary; USER ID=sa; PASSWORD=qazxsw; "
		Str_Connect2000= "PROVIDER=SQLOLEDB; DATA SOURCE=192.168.12.100,1433; INITIAL CATALOG=weblibrary; USER ID=sa; PASSWORD=qazxsw; "
		ConSql2000.open Str_Connect2000
		
		
		
	'------------------------------------------------------------------------Calendar QS----------------------------------------------------------------------------------------------	
	    Set ConActivity = Server.CreateObject("ADODB.Connection")
		'Str_Connect = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source ="&Server.mappath("\030_KM\_db\KM.mdb")
		Str_ConnectActivity = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source=E:\030_KM\_block\qos\db\BookMeeting.mdb"
		'Str_Connect = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source="&server.MapPath("_db/tpd.mdb")
		'Str_Connect = "Provider=Microsoft.Jet.OLEDB.4.0;Jet Oledb:Database Password=coolooc;Data Source=C:\inetpub\wwwroot\weblib_admin\_db\tpd.mdb"
    	
		'Str_Connect = "Provider=SQLOLEDB;Server=(local);integrated security=SSPI;uid=sa;pwd=coolooc;database=library"
		'Str_Connect = "driver={SQL Server};server=localhost;uid=sa;pwd=coolooc;database=library;" 
		ConActivity.open Str_ConnectActivity

		Function GetSingleField(Table,Field,Condition)
		'on error resume next
			Set rs_g = ConActivity.Execute("SELECT "&Field&" FROM "&Table&" "&Condition)
		
					If Not rs_g.Eof Then
						GetSingleField = rs_g(0)
					Else 
						GetSingleField = 0
					End If
		
			rs_g.Close()
		End Function
			
		Function GetMax(Table,Field,Condition)
			Set rs_g = ConActivity.Execute("SELECT Max("&Field&") FROM "&Table&" "&Condition)
		
					If Not rs_g.Eof Then
						GetMax = rs_g(0)
					Else 
						GetMax = 0
					End If
		
			rs_g.Close()
		End Function
		
		Function GetMin(Table,Field,Condition)
			Set rs_g = ConActivity.Execute("SELECT Min("&Field&") FROM "&Table&" "&Condition)
		
					If Not rs_g.Eof Then
						GetMin = rs_g(0)
					Else 
						GetMin = 0
					End If
		
			rs_g.Close()
		End Function
		
		
		Function GetBetweenTime(startDate,endDate)
			Set rs_g = ConActivity.Execute("SELECT count(B_Id) as Bid from Tb_Book where (Hour(B_TimeStart) Between "&startDate&" and "&endDate&" and Minute(B_TimeStart) Between 0 and 59) and (Hour(B_TimeEnd) Between "&startDate&" and "&endDate&" and Minute(B_TimeStart) Between 0 and 59 ) ")
		
					If Not rs_g.Eof Then
						GetBetweenTime = rs_g(0)
					Else 
						GetBetweenTime = 0
					End If
		
			rs_g.Close()
		End Function
		
		function getDataCalendarActivity(GDate)
			SQL = "Select count(A_ID) from Tb_Activity where  A_Flag = true  and  A_StartDate <= #"&GDate&"# and A_EndDate >= #"&GDate&"#"
			set RecData = Server.CreateObject("ADODB.RECORDSET")
			RecData.open SQL,ConActivity,1,3
			gd = RecData.RecordCount
			If Not RecData.Eof Then
				getDataCalendarActivity = RecData(0)
			Else 
				getDataCalendarActivity = 0
			End If		
			RecData.Close()
		end function
		function getDataCalendarBooking(GDate)
			SQL = "Select count(B_ID) from Tb_Book where B_Flag = True and  B_StartDate <= #"&GDate&"# and B_EndDate >= #"&GDate&"#"
			set RecData = Server.CreateObject("ADODB.RECORDSET")
			RecData.open SQL,ConActivity,1,3
			gd = RecData.RecordCount
			If Not RecData.Eof Then
				getDataCalendarBooking = RecData(0)
			Else 
				getDataCalendarBooking = 0
			End If		
			RecData.Close()
		end function
   '-------------------------------------------------------------------------end calendar QS------------------------------------------------------------------------------------------		
		
		
		Dim Tab , arrAge , arrDuring , arrSex , arrTemplate , arrOrderType , arrFName , arrPrefer , arrNumMonth , arrTxtMonth , Arr_Bgcolor
		'arrSearchType=Array("","Travel Service","Education Service")
		arrTime=Array("","0:00","0:30","1:00","1:30 ","2:00 ","2:30 ","3:00 ","3:30 ","4:00 ","4:30 ","5:00 ","5:30 ","6:00 ","6:30 ","7:00 ","7:30 ","8:00 ","8:30 ","9:00 ","9:30 ","10:00 ","10:30 ","11:00 ","11:30 ","12:00 ","12:30 ","13:00 ","13:30 ","14:00 ","14:30 ","15:00 ","15:30 ","16:00 ","16:30 ","17:00 ","17:30 ","18:00 ","18:30 ","19:00 ","19:30 ","20:00 ","20:30 ","21:00 ","21:30 ","22:00 ","22:30 ","23:00 ","23:30")
		arrAge = Array("","1","2","3","4","5","6","7","8","9","10","11","12","13","14","15","16","17","18","19","20","21","22","23","24","25","26","27","28","29","30","31","32","33","34","35","36","37","38","39","40","41","42","43","44","45","46","47","48","49","50","51","52","53","54","55","56","57","58","59","60","61","62","63","64","65","66","67","68","69","70","71","72","73","74","75","76","77","78","79","80","81","82","83","84","85","86","87","88","89","90")
		arrDuring = Array("","1","2","3","4","5","6","7","8","9","10")
		arrSex = Array("","Male","Female")
		arrTemplate = Array("","1","2")
		arrOrderType = Array("","Hotel")
		arrPrefer = Array("","AISLE","WINDOW","ANY")
		arrNumMonth = Array("","01","02","03","04","05","06","07","08","09","10","11","12")
		thmonth=array("","ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค.")
		enmonth = array("","Jan","Feb","Mar","Apr","May","June","July","Aug","Sep","Oct","Nov","Dec")
		thmonthFull=array("","มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน","กรกฏาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม")
		enmonthFull = array("","January","February","March","April","May","June","July","August","September","October","November","December")
		thaiday = Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
		Engday = Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
		Arr_m = Array("","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
		Arr_mn = Array("","1","2","3","4","5","6","7","8","9","10","11","12")
		arrTxtMonth = Array("","Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
		Arr_Bgcolor = Array("#EEEEDD","#F3F3F8","#DDF4FF","#FFF5EC","#EEEEDD")
		ArrPicture = Array("Pic1","Pic2","Pic3","Pic4","Pic5","Pic6","Pic7")
		ArrPicturePackage = Array("PgImage","PgImage2","PgImage3")
		ArrPictureEducation = Array("PgImage","PgImage2","PgImage3")
		arrClass=Array("","*","**","***","****","*****")
		Tab = "&nbsp;&nbsp;&nbsp;"
		TabMenu = "&nbsp;&nbsp;&nbsp;&nbsp;"
		height_outer = "100%"
		height_inner = "100%"
		bg_head =" background='"&path_link&"_images/bg_table.gif' height='25' "
		bg_head1 =" bgcolor='#f7f7f7' "
		border_head = "bordercolor=#F0F0F0"
		bg_head2=" bgColor=#A7C9F0 "
		bg_table_admin="#006600"
		
		'department=Array("","เลขาธิการฯ","รองเลขาธิการฯ","นวช. 10 ชช.","เภสัชกร 9 วช.","นวช. 9 ชช.","นิติกร 9 ชช.","ผู้ช่วยเลขาธิการฯ","สำนักควบคุมเครื่องสำอางและวัตถุอันตราย","สำนักงานเลขานุการกรม","กองควบคุมเครื่องมือแพทย์","กองควบคุมยา","กองควบคุมวัตถุเสพติด","กองควบคุมอาหาร",".สำนักด่านอาหารและยา","กองแผนงานและวิชาการ","กองพัฒนาศักภาพผู้บริโภค","ศพช.","กลุ่มควบคุมเครื่องสำอาง","กลุ่มควบคุมวัตถุอันตราย","กลุ่มกฎหมายอาหารและยา","กลุ่มตรวจสอบภายใน","กอง คบ.","ตส.","ศูนย์เทคโนโลยีสารสนเทศ","กลุ่มพัฒนาระบบบริหาร","กลุ่มผลิตภัณฑ์ทางเลือกเพื่อสุขภาพ","IPCS","สสจ.","ศูนย์บริการผลิตภัณฑ์สุขภาพเบ็ดเสร็จ")
				department=Array("","เลขาธิการฯ","รองเลขาธิการฯ","นวช. 10 ชช.","เภสัชกร 9 วช.","นวช. 9 ชช.","นิติกร 9 ชช.","ผู้ช่วยเลขาธิการฯ","สำนักควบคุมเครื่องสำอางและวัตถุอันตราย","สำนักงานเลขานุการกรม","กองควบคุมเครื่องมือแพทย์","สำนักยา","กองควบคุมวัตถุเสพติด","สำนักอาหาร","สำนักด่านอาหารและยา","กองแผนงานและวิชาการ","กองพัฒนาศักภาพผู้บริโภค","ศพช.","กลุ่มควบคุมเครื่องสำอาง","กลุ่มควบคุมวัตถุอันตราย","กลุ่มกฎหมายอาหารและยา","กลุ่มตรวจสอบภายใน","กอง คบ.","ตส.","ศูนย์เทคโนโลยีสารสนเทศ","กลุ่มพัฒนาระบบบริหาร","กลุ่มผลิตภัณฑ์ทางเลือกเพื่อสุขภาพ","สำนักความร่วมมือระหว่างประเทศ","สสจ.","ศูนย์บริการผลิตภัณฑ์สุขภาพเบ็ดเสร็จ")

		query_string = Request.ServerVariables("query_string")
		Pathdraw = "../img_drawing/"
		PathdrawRent = "../../img_drawingRent/"
		PathPackage = "../../_ImgUpload/img_package/"
		PathPackage_Front = "_ImgUpload/img_package/"
		PathEducation= "../../_ImgUpload/img_Education/"
		PathEducation_Front= "_ImgUpload/img_Education/"
		PathCarrental = "../../Carrental/"
		PathEx = "../../_ImgUpload/Img_Hotel/"
		PathEx_Front = "_ImgUpload/Img_Hotel/"
		Pathallpicture 	= "../../_ImgUpload/Img_allpicture/"
		Pathallpicture_Front= "_ImgUpload/Img_allpicture/"
		PathGallery="../../_ImgUpload/Img_gallery/"
		PathGallery_Front="_ImgUpload/Img_gallery/"
		PathWebboard="../../_ImgUpload/Img_webboard/"

ID_CopGroup="4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40,41,42,43,44,45,46,47,48,49,50"



'==============Set Size In PageInfo==================
Select Case request("pagename")
Case "Package A","Package B","Package C","Package D":sizeeditor="162"
Case "Package A Detail","Package B Detail","Package C Detail","Package D Detail":sizeeditor="420"
Case "Short Our Staff":sizeeditor="155"
Case "Short Brief History":sizeeditor="150"
Case "Short Our Business":sizeeditor="400"
Case "Contact Us":sizeeditor="380"
Case "Hotel Reservations":sizeeditor="560"
Case Else :sizeeditor="740"
End Select 
'=================================================




'======================TH-EN====================================
if session("lang")="" then session("lang")="TH"
if request("lang")<>"" then session("lang")=request("lang")
lang=session("lang")
lang2=replace(lang,"TH","")


room_=SwapWord("ห้อง","Rooms",lang)
'================================================================
		
'========================Start Function for QS system========================================
function getDepartmentname(getHidDid)
	sql="select D_Name from Tb_Department where D_Id='"&getHidDid&"'"
	set rec = Server.CreateObject("ADODB.RECORDSET")
	rec.open sql,ConQS,1,3
	while  not rec.EOF
	Rreturn = rec("D_Name")
	rec.MoveNext
	wend
	getDepartmentname = Rreturn
end function
function getJoinDepartmentname(getMId)
	sql="select Tb_Department.D_Name from Tb_Department  inner join Tb_Manual on  Tb_Department.D_Id=Tb_Manual.D_Id where M_Id="&getMId
	set rec = Server.CreateObject("ADODB.RECORDSET")
	rec.open sql,ConQS,1,3
	while  not rec.EOF
	Rreturn = rec("D_Name")
	rec.MoveNext
	wend
	getJoinDepartmentname = Rreturn
end function
function getJoinDepartmentId(getMId)
	sql="select Tb_Department.D_Id from Tb_Department  inner join Tb_Manual on  Tb_Department.D_Id=Tb_Manual.D_Id where M_Id="&getMId
	set rec = Server.CreateObject("ADODB.RECORDSET")
	rec.open sql,ConQS,1,3
	while  not rec.EOF
	Rreturn = rec("D_Id")
	rec.MoveNext
	wend
	getJoinDepartmentId = Rreturn
end function
Function GetSingleFieldQS(Table,Field,Condition)
'on error resume next
	Set rs_g = ConQS.Execute("SELECT "&Field&" FROM "&Table&" "&Condition)

			If Not rs_g.Eof Then
				GetSingleFieldQS = rs_g(0)
			Else 
				GetSingleFieldQS = 0
			End If

	rs_g.Close()
End Function
Function GetCountRowQS(Table,Field,Condition)
'on error resume next
	Set rs_g = ConQS.Execute("SELECT count("&Field&") FROM "&Table&" "&Condition)
	If Not rs_g.Eof Then
				GetCountRowQS = rs_g(0)
			Else 
				GetCountRowQS = 0
			End If		
	rs_g.Close()
End Function
Sub AddCarToTable(Field1,Field2)
	Set rs_g = ConQS.Execute("insert into  Tb_RunNumCAR (CAR_ID,M_Code) values ('"&Field1&"','"&Field2&"')")	
	rs_g.Close()
end Sub
Function checkInternalAuditData(mCode,auLevel,auYear)
			
			sql = "select count(ID) from Tb_Internalaudit where M_Code='"&mCode&"' and Audit_Level='"&auLevel&"' and Audit_Year='"&auYear&"' "
			
			Set rs_g = ConQS.Execute(sql)

			If Not rs_g.Eof Then
				checkInternalAuditData = rs_g(0)
			Else 
				checkInternalAuditData = 0
			End If

	rs_g.Close()
End Function


Function CheckLevelSuccess1(DId)
	set recAnalisProcedure = Server.CreateObject("ADODB.RECORDSET")
	sql_chkAnalis = "SELECT Tb_AnalisProcedure.M_Code FROM Tb_AnalisProcedure INNER JOIN Tb_Manual ON Tb_AnalisProcedure.M_Id = Tb_Manual.M_Id where Tb_Manual.D_Id='"&DId&"' "
	recAnalisProcedure.open sql_chkAnalis,ConQS,1,3
	getCountRow = recAnalisProcedure.RecordCount
	CheckLevelSuccess1 = getCountRow
	recAnalisProcedure.Close()
End function


Function CheckLevelSuccess2(DId)
	set recAnalisProcedure = Server.CreateObject("ADODB.RECORDSET")
	sql_chkAnalis = "SELECT * from Tb_Review where D_Id='"&DId&"' "
	recAnalisProcedure.open sql_chkAnalis,ConQS,1,3
	getCountRow = recAnalisProcedure.RecordCount
	CheckLevelSuccess2 = getCountRow
	recAnalisProcedure.Close()
End Function

Function CheckLevelSuccess3(DId)
	set recAnalisProcedure = Server.CreateObject("ADODB.RECORDSET")
	sql_chkAnalis = "SELECT * from Tb_InternalAudit where Audit_Depart='"&DId&"' "
	recAnalisProcedure.open sql_chkAnalis,ConQS,1,3
	getCountRow = recAnalisProcedure.RecordCount
	CheckLevelSuccess3 = getCountRow
	recAnalisProcedure.Close()
End Function

Function CheckLevelSuccess4(DId)
	set recAnalisProcedure = Server.CreateObject("ADODB.RECORDSET")
	sql_chkAnalis = "SELECT * FROM Tb_ManagementReview where D_Id='"&DId&"' "
	recAnalisProcedure.open sql_chkAnalis,ConQS,1,3
	getCountRow = recAnalisProcedure.RecordCount
	CheckLevelSuccess4= getCountRow
	recAnalisProcedure.Close()
End Function


Function  getCARNumber(Did,M_Code,AuditLevel)
	set rec = Server.CreateObject("ADODB.RECORDSET")
	dim getData
	sql = "select  No_Car_Par from Tb_InternalAudit where Audit_Depart='"&Did&"' and M_Code='"&M_Code&"' and  Audit_Level='"&AuditLevel&"' and Audit_DocType='NC' "
	rec.open sql,ConQS,1,3
	while not rec.EOF
		getData = getData&rec("No_Car_par")&" <br>"
	rec.MoveNext
	Wend
	rec.Close()
	getCARNumber = getData
End Function


Function  getPARNumber(Did,M_Code,AuditLevel)
	set rec = Server.CreateObject("ADODB.RECORDSET")
	dim getData
	sql = "select  No_Car_Par from Tb_InternalAudit where Audit_Depart='"&Did&"' and M_Code='"&M_Code&"' and  Audit_Level='"&AuditLevel&"' and Audit_DocType='OBS' "
	rec.open sql,ConQS,1,3
	while not rec.EOF
		getData = getData&rec("No_Car_par")&" <br>"
	rec.MoveNext
	Wend
	rec.Close()
	getPARNumber = getData
End Function

Function getCountRowAnalis(Did)
	set rec = Server.CreateObject("ADODB.RECORDSET")
	dim getData
	sql = "select count(Tb_AnalisProcedure.M_Id) as M_Id  from (Tb_AnalisProcedure inner join Tb_Manual on Tb_AnalisProcedure.M_Id=Tb_Manual.M_Id) inner join Tb_Department on Tb_Manual.D_Id = Tb_Department.D_Id where Tb_Department.D_Id='"&Did&"' and Tb_Manual.M_Main=1 and Tb_Manual.M_Reserve=0" 
	rec.open sql,ConQS,1,3
	getData = rec("M_Id")
	rec.Close()
	getCountRowAnalis = getData
End Function

Function getPermission(getMemberID,getField)
	Set result = Server.CreateObject("ADODB.RECORDSET")
	SQL = "select  "&getField&"  from Tb_Level where L_Email = '"&getMemberID&"' "
	 result.open SQL,ConQS,1,3
		If not result.Eof then
			getPermission = result(0)
		else
			getPermission = 0
		End if
	result.Close()
End Function

'=========================End Function for QS system========================================
				
			


	Sub Sub_Link(str_link,str_field,BYref str_order,BYref con_order)
		IF CStr(str_order) = CStr(str_field) Then
			IF con_order = "Desc" Then
				con_order = "Asc"
				str_img ="<img src='../image/arrowdown2.gif' border='0'>"
			Else
				con_order = "Desc"
				str_img ="<img src='../image/arrowup2.gif' border='0'>"
			End IF
		Else
			str_img ="<img src='../image/arrowup2.gif' border='0'>"
		End IF
		response.write "<a href='?str_order="&str_field&"&con_order="&con_order&"'>"&str_link&str_img&"</a>"
	End Sub
	
			Function RegExps(strMatchPattern, strPhrase)
				Dim objRegEx, Match, Matches, StrReturnStr
				Set objRegEx = New RegExp 
				objRegEx.Global = True
				objRegEx.IgnoreCase = True
				objRegEx.Pattern = strMatchPattern 
				Set Matches = objRegEx.Execute(strPhrase) 
				found = 0
				For Each Match in Matches
				found = found +1
				Next
				RegExps=found
			End function
			path_info = Request.ServerVariables("path_info")
			num_path = RegExps("/",path_info)
		'==================server  (num_path - 1)==============
			for p = 1 to (num_path - 2)
				path_link = path_link &"../"
			next
			'background =""&path_link&"images/bg_main.gif"
			'================
			 
			Sub ShowNew(str_date)
				IF str_date = Date or str_date > (Date -3) Then
					response.write "<img src='"&path_link&"admin/image/new_pic.gif'>"
				End IF
			End Sub
			
			Function ReplaceData(Data,NewData)
				ReplaceData = replace(Data,NewData,"<b><font color='#ff0000'>"&NewData&"</font></b>")
			End Function
			

'============Function Replace \n			
			Function ReplaceSn(NewData)
			IF NewData <> "" Then
				ReplaceSn = Replace(NewData,vbCrLf,"\n") 
			Else
				ReplaceSn	= "-"
			End IF
			End Function

			Function ChangetoYesNo(val)
				IF val = "on" or val = "1" Then
					ChangetoYesNo = True
				Else
					ChangetoYesNo = False
				End IF
			End Function
			
			Sub ShowStar(str_num)
				
				sp_star = split(str_num,".")
				on error resume next
				if Ubound(sp_star) > 0 Then
					str_star1 = sp_star(1)
					str_star0 = sp_star(0)
				Else
					str_star0 = sp_star(0)
				End IF
					
				for i = 1 to str_star0
					response.write "<img src='"&path_link&"IMAGE/starfull.gif'>&nbsp;"
				Next
				IF str_star1 <> "" Then
					response.write "<img src='"&path_link&"IMAGE/starhalf.gif'>&nbsp;"
				End IF
			End Sub
'================Function Multi Search 
			Function MultiSearch(Data,Fieldsearch)
				IF Data <> "" THen
					sp_data = split(Data," ")
					sp_field = split(Fieldsearch,",")
					For i = 0 To Ubound(sp_field)
						For j = 0 To Ubound(sp_data)
						QuerySearch = QuerySearch&sp_field(i) &" like '%"&sp_data(j)&"%' or " 
						Next
					Next
					MultiSearch = CutData(QuerySearch,3)
				End IF		
			End Function
'================ End Function 
'============== Function Replace Multi Value==============
Function ReplaceMulti(Text,Data)
		IF Data <> "" and Text <> "" Then
			sp_data = split(Data," ")
			For j = 0 To Ubound(sp_data)
				n_data = "<font class=r>"&sp_data(j)&"</font>"
				Text = replace(Text,sp_data(j),n_data)
			Next
			ReplaceMulti = Text
		Else
			ReplaceMulti = Text
		End IF
End Function
'===========Function Cut Data 
			Function CutData(Data,num)
				IF Data = "" Then CutData = Data Else CutData = left(Data,len(Data)-num) End IF
			End Function

'===========Function Change Blank To "-"
			Function Str_blank(Data)
				IF Data = "" or isnull(Data) Then Str_blank = "-" Else Str_blank = Data End IF
			End Function
'=========== Function cut character
			Function CutChar(Data,num)
				IF len(Data) <= num Then
					CutChar =	Data
				Else
					CutChar = left(Data,num)&"....."
				End IF
			End Function
'============ Function Clear Single Quote 
			Function ClearSingleQuote(Data)
				ClearSingleQuote = trim(replace(Data,"'",""))
			End Function


'=========== Function Change To Format Code
			Function changetocode(val)
				IF val <= 9   Then
					str_val = "0000"&val
				ElseIF val >9 and val <= 99 Then
					str_val = "000"&val
				ElseIF val >99 and  val <= 999 Then
					str_val = "00"&val
				ElseIF val >999 and val <= 9999 Then
 					str_val = "0"&val
				Else
					str_val = val
				End IF
				changetocode = str_val
			End Function 




			Sub Head()
				With Response
				.Write "<html>"&_
				"<head>"&_
				"<title>"&website&""&_
				"</title>"&_
				"</head>"&_
				"<style type=""text/css"">"&_
					"A:link { Color:#0000FF; text-decoration: none }"&_
					"A:visited { Color:#0000FF; text-decoration: none }"&_
					"A:hover { Color:#000000; text-decoration: none }"&_
				"</style>"
				End With
			End Sub
	
	
	
	
			Function Change_date(rsDate,lang)
				
					thmonth=array("","ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค.")
					enmonth = array("","Jan","Feb","Mar","Apr","May","June","July","Aug","Sep","Oct","Nov","Dec")
					mday =day(rsDate)
					if lang = "TH" Then
						mmonth = thmonth(month(rsDate))
					Else
						mmonth = enmonth(month(rsDate))
					End IF	
					myear = year(rsDate) +543
					Change_date =  mday&" "&mmonth&" "&myear
			
			End Function	
	
Function  SwapWord(wordTH,wordEN,lang)
	Select Case lang
	Case "TH":SwapWord=wordTH
	Case "EN":SwapWord=wordEN
	End Select
End Function
	

	Sub MainTopic(Text,Color)
%>
<TABLE cellSpacing=0 cellPadding=0 width="100%"  border=0 class="text1" background="<%=color%>" ><TBODY><TR>
<TD height="31" align=left background="<%=path_link%>images/bg_header.gif" width="580"> 
<b>&nbsp;<font color="#FFFFFF"><%=TabMenu&text%></font></b></TD></TR></TBODY></TABLE>
<%
	End Sub


Function textToImg(data,text,pathImg)
	IF data="" or isnull(data) then
	textToImg="-"
	Else
		IF UCASE(text)=text Then
			for u=1 to len(data)
			textToImg=textToImg&"<img src="&pathImg&" border=0 aling='absmiddle'>"
			next
		End if
	End IF
End Function



Sub FormSetHotel%>
<table width="85%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td  background="<%=path_link%>_images/box/box-c-left.gif" width="18" height="12"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td   background="<%=path_link%>_images/box/box-c-center.gif"  height="12"><img src="<%=path_link%>_images/SPACER.GIF"  height="12"></td>
                                    <td   background="<%=path_link%>_images/box/box-c-right.gif" width="18" height="12"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                  </tr>
                                  <tr> 
                                    
    <td width="18"  valign="top" background="<%=path_link%>_images/box/box-line-left.gif"><img src="<%=path_link%>_images/SPACER.GIF" width="18" height="30"></td>
                                    
                
    <td valign="top" width="100%" bgcolor="#FFFFFF"><%HotelTypeId=request("HotelTypeId")
SubHotelTypeId=request("SubHotelTypeId")
ThirdHotelTypeId=request("ThirdHotelTypeId")%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="text1" align="center">

<script language="JavaScript">

	function chk_adddisable0(id)
	{
			f = document.form;
			fa = window;
			f.SubHotelTypeId.value="";
			f.ThirdHotelTypeId.value="";
			sp_id = f.allgroup.value.split(",");
			for(i=0;i<sp_id.length;i++)
			{
				if(sp_id[i] != "") 
				{eval("f.HotelType"+sp_id[i]+".value=''");
					if(sp_id[i] == id)
					{eval("f.HotelType"+sp_id[i]+".style.display=''");
					eval("fa.HotelType"+sp_id[i]+".style.display=''");}
					else
					{eval("fa.HotelType"+sp_id[i]+".style.display='none'");}}
			}
			
			spsub_id = f.allsubgroup.value.split(",");
			for(i=0;i<spsub_id.length;i++)
			{
				if(spsub_id[i] != "") 
				{eval("fa.SubHotelType"+spsub_id[i]+".style.display='none'");}
			}
		}



	function chk_adddisable(id)
	{
			f = document.form;
			fa = window;
			f.ThirdHotelTypeId.value="";
			spsub_id = f.allsubgroup.value.split(",");
			for(i=0;i<spsub_id.length;i++)
			{
				if(spsub_id[i] != "") 
				{eval("f.SubHotelType"+spsub_id[i]+".value=''");
					if(spsub_id[i] == id)
					{eval("f.SubHotelType"+spsub_id[i]+".style.display=''");
					eval("fa.SubHotelType"+spsub_id[i]+".style.display=''");}
					else
					{eval("fa.SubHotelType"+spsub_id[i]+".style.display='none'");}}
			}
		}
		
		
		

</script>
                            <tr> 
                              
          <td width="380">First Hotel Type </td>
            <td width="200"> <%
		Table2= "TabHotelType"
		Condition2= ""
		SelectName= "HotelTypeId"
		FieldValue= "HotelTypeId"
		FieldDesc= "HotelType"
		ValueExp=HotelTypeId
		CJava= "onchange='chk_adddisable0(this.value)'"
		Str_Property= ""
		 Call ListBox("=========== All Hotel Type ===========",Table2,Condition2,SelectName,FieldValue,FieldDesc,ValueExp,CJava,Str_Property)
%> </td>
          </tr>
          <% 
		  Set rs = server.createobject("Adodb.recordset")
			rs.open "select * from TabHotelType ",con,1,1
			while not rs.eof
			if countrecord("TabSubHotelType","Where HotelTypeId = "&rs("HotelTypeId")) > 0 Then
			allgroup = allgroup&rs("HotelTypeId")&","
			IF CStr(HotelTypeId) = CStr(rs("HotelTypeId"))	Then
				SubHotelType_Style = ""
			Else
				SubHotelType_Style=" style='display:none'"
			End IF
			%>
                            <tr id="HotelType<%=rs("HotelTypeId")%>" <%=SubHotelType_Style%>> 
                              
          <td height="25"   >Second Hotel Type<strong>&nbsp;&nbsp;</strong></td>
            <td ><select name="HotelType<%=rs("HotelTypeId")%>" class="textbox"  onchange="chk_adddisable(this.value);document.form.SubHotelTypeId.value=this.value">
                <option value="">==== All Second Hotel Type ====
                <%
				Set rs_sub = server.createobject("Adodb.recordset")
				rs_sub.open "select * from TabSubHotelType Where HotelTypeId = "&rs("HotelTypeId"),con,1,1
				while not rs_sub.eof
				IF CStr(SubHotelTypeId) = CStr(rs_sub("SubHotelTypeId")) Then
					sub_sel = " selected "
				Else
					sub_sel = ""
				End IF
			%>
                <option value="<%=rs_sub("SubHotelTypeId")%>" <%=sub_sel%>><%=rs_sub("SubHotelType")%> 
                <%		
				rs_sub.movenext
				wend
			%>
              </select> </td>
          </tr>
          <%	
			end if
				rs.movenext
				wend
				
				Call Closerecord(rs)
				%>
				
				
				
          <% 
		  Set rs = server.createobject("Adodb.recordset")
			rs.open "select * from TabSubHotelType",con,1,1
			while not rs.eof
			if countrecord("TabThirdHotelType","Where SubHotelTypeId = "&rs("SubHotelTypeId")) > 0 Then
			allsubgroup = allsubgroup&rs("SubHotelTypeId")&","
			IF CStr(SubHotelTypeId) = CStr(rs("SubHotelTypeId"))	Then
				ThirdHotelType_Style = ""
			Else
				ThirdHotelType_Style=" style='display:none'"
			End IF
			%>
                            <tr id="SubHotelType<%=rs("SubHotelTypeId")%>" <%=ThirdHotelType_Style%>> 
                              <td   >Third Hotel Type<strong>&nbsp;&nbsp;</strong></td>
            <td ><select name="SubHotelType<%=rs("SubHotelTypeId")%>" class="textbox"  onchange="document.form.ThirdHotelTypeId.value=this.value">
                <option value="">==== All Third Hotel Type =====
                <%
				Set rs_sub = server.createobject("Adodb.recordset")
				rs_sub.open "select * from TabThirdHotelType Where SubHotelTypeId = "&rs("SubHotelTypeId"),con,1,1
				while not rs_sub.eof
				IF CStr(ThirdHotelTypeId) = CStr(rs_sub("ThirdHotelTypeId")) Then
					sub_sel = " selected "
				Else
					sub_sel = ""
				End IF
			%>
                <option value="<%=rs_sub("ThirdHotelTypeId")%>" <%=sub_sel%>><%=rs_sub("ThirdHotelType")%> 
                <%		
				rs_sub.movenext
				wend
			%>
              </select> </td>
          </tr>
          <%	
			end if
				rs.movenext
				wend
				
				Call Closerecord(rs)
				%>

<tr>
                                <td>&nbsp;</td>
								<td><input type="image" src="<%=path_link%>_images/button/bt_searchnew.gif" class="textbox"  value="Search" ></td>
                              </tr>
						<input type="hidden" name="SubHotelTypeId" value="<%=SubHotelTypeId%>"> 
						<input type="hidden" name="ThirdHotelTypeId" value="<%=ThirdHotelTypeId%>">
						<input type="hidden" name="allgroup" value="<%=allgroup%>">
						<input type="hidden" name="allsubgroup" value="<%=allsubgroup%>">

</table></td>
                                    <td width="18"  valign="top"  background="<%=path_link%>_images/box/box-line-right.gif"><img src="<%=path_link%>_images/SPACER.GIF" width="18" height="30"></td>
                                  </tr>
                                  <tr> 
                                    <td width="18" height="12" valign="top"  background="<%=path_link%>_images/box/box-c-below-left.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td valign="top" background="<%=path_link%>_images/box/box-c-below-center.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td valign="top" background="<%=path_link%>_images/box/box-c-below-right.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                  </tr>
                                </table>

<%End Sub


Sub FormSearchHotel%>
<table width="10" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td  background="<%=path_link%>_images/box/box-c-left.gif" width="18" height="12"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td   background="<%=path_link%>_images/box/box-c-center.gif"  height="12"><img src="<%=path_link%>_images/SPACER.GIF"  height="12"></td>
                                    <td   background="<%=path_link%>_images/box/box-c-right.gif" width="18" height="12"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                  </tr>
                                  <tr> 
                                    
    <td width="18"  valign="top" background="<%=path_link%>_images/box/box-line-left.gif"><img src="<%=path_link%>_images/SPACER.GIF" width="18" height="30"></td>
                                    
                
    <td valign="top" width="100%" bgcolor="#FFFFFF"><%HotelTypeId=request("HotelTypeId")
SubHotelTypeId=request("SubHotelTypeId")
ThirdHotelTypeId=request("ThirdHotelTypeId")%>
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="text1" align="center">
<form name=form>
<script language="JavaScript">

	function chk_adddisable0(id)
	{
			f = document.form;
			fa = window;
			f.SubHotelTypeId.value="";
			f.ThirdHotelTypeId.value="";
			sp_id = f.allgroup.value.split(",");
			for(i=0;i<sp_id.length;i++)
			{
				if(sp_id[i] != "") 
				{eval("f.HotelType"+sp_id[i]+".value=''");
					if(sp_id[i] == id)
					{eval("f.HotelType"+sp_id[i]+".style.display=''");
					eval("fa.HotelType"+sp_id[i]+".style.display=''");}
					else
					{eval("fa.HotelType"+sp_id[i]+".style.display='none'");}}
			}
			
			spsub_id = f.allsubgroup.value.split(",");
			for(i=0;i<spsub_id.length;i++)
			{
				if(spsub_id[i] != "") 
				{eval("fa.SubHotelType"+spsub_id[i]+".style.display='none'");}
			}
		}



	function chk_adddisable(id)
	{
			f = document.form;
			fa = window;
			f.ThirdHotelTypeId.value="";
			spsub_id = f.allsubgroup.value.split(",");
			for(i=0;i<spsub_id.length;i++)
			{
				if(spsub_id[i] != "") 
				{eval("f.SubHotelType"+spsub_id[i]+".value=''");
					if(spsub_id[i] == id)
					{eval("f.SubHotelType"+spsub_id[i]+".style.display=''");
					eval("fa.SubHotelType"+spsub_id[i]+".style.display=''");}
					else
					{eval("fa.SubHotelType"+spsub_id[i]+".style.display='none'");}}
			}
		}
		
		
		

</script>
                            <tr> 
                              
            <td width="120" style="display:none">First Hotel Type </td>
            <td> <%
		Table2= "TabHotelType"
		Condition2= " Order By Numberlist"
		SelectName= "HotelTypeId"
		FieldValue= "HotelTypeId"
		FieldDesc= "HotelType"
		ValueExp=HotelTypeId
		CJava= "onchange='chk_adddisable0(this.value)'"
		Str_Property= ""
		 Call ListBox("=== All First Hotel Type ===",Table2,Condition2,SelectName,FieldValue,FieldDesc,ValueExp,CJava,Str_Property)
%> </td>
          </tr>
          <% 
		  Set rs = server.createobject("Adodb.recordset")
			rs.open "select * from TabHotelType",con,1,1
			while not rs.eof
			if countrecord("TabSubHotelType","Where HotelTypeId = "&rs("HotelTypeId")) > 0 Then
			allgroup = allgroup&rs("HotelTypeId")&","
			IF CStr(HotelTypeId) = CStr(rs("HotelTypeId"))	Then
				SubHotelType_Style = ""
			Else
				SubHotelType_Style=" style='display:none'"
			End IF
			%>
                            <tr id="HotelType<%=rs("HotelTypeId")%>" <%=SubHotelType_Style%>> 
                              <td   style="display:none">Second Hotel Type<strong>&nbsp;&nbsp;</strong></td>
            <td ><select name="HotelType<%=rs("HotelTypeId")%>" class="textbox"  onchange="chk_adddisable(this.value);document.form.SubHotelTypeId.value=this.value">
                <option value="">== All Second Hotel Type ==
                <%
				Set rs_sub = server.createobject("Adodb.recordset")
				rs_sub.open "select * from TabSubHotelType Where HotelTypeId = "&rs("HotelTypeId")&" Order By Numberlist",con,1,1
				while not rs_sub.eof
				IF CStr(SubHotelTypeId) = CStr(rs_sub("SubHotelTypeId")) Then
					sub_sel = " selected "
				Else
					sub_sel = ""
				End IF
			%>
                <option value="<%=rs_sub("SubHotelTypeId")%>" <%=sub_sel%>><%=rs_sub("SubHotelType")%> 
                <%		
				rs_sub.movenext
				wend
			%>
              </select> </td>
          </tr>
          <%	
			end if
				rs.movenext
				wend
				
				Call Closerecord(rs)
				%>
				
				
				
          <% 
		  Set rs = server.createobject("Adodb.recordset")
			rs.open "select * from TabSubHotelType",con,1,1
			while not rs.eof
			if countrecord("TabThirdHotelType","Where SubHotelTypeId = "&rs("SubHotelTypeId")) > 0 Then
			allsubgroup = allsubgroup&rs("SubHotelTypeId")&","
			IF CStr(SubHotelTypeId) = CStr(rs("SubHotelTypeId"))	Then
				ThirdHotelType_Style = ""
			Else
				ThirdHotelType_Style=" style='display:none'"
			End IF
			%>
                            <tr id="SubHotelType<%=rs("SubHotelTypeId")%>" <%=ThirdHotelType_Style%>> 
                              <td   style="display:none">Third Hotel Type<strong>&nbsp;&nbsp;</strong></td>
            <td ><select name="SubHotelType<%=rs("SubHotelTypeId")%>" class="textbox"  onchange="document.form.ThirdHotelTypeId.value=this.value">
                <option value="">=== All Third Hotel Type===</option>
                <%
				Set rs_sub = server.createobject("Adodb.recordset")
				rs_sub.open "select * from TabThirdHotelType Where SubHotelTypeId = "&rs("SubHotelTypeId")&" Order By ThirdHotelType",con,1,1
				while not rs_sub.eof
				IF CStr(ThirdHotelTypeId) = CStr(rs_sub("ThirdHotelTypeId")) Then
					sub_sel = " selected "
				Else
					sub_sel = ""
				End IF
			%>
                <option value="<%=rs_sub("ThirdHotelTypeId")%>" <%=sub_sel%>><%=rs_sub("ThirdHotelType")%> 
                <%		
				rs_sub.movenext
				wend
			%>
              </select> </td>
          </tr>
          <%	
			end if
				rs.movenext
				wend
				
				Call Closerecord(rs)
				%>

<tr>
                                <td style="display:none">&nbsp;</td>
								<td><input type="image" src="<%=path_link%>_images/button/bt_searchnew.gif" class="textbox"  value="Search"></td>
                              </tr>
						<input type="hidden" name="SubHotelTypeId" value="<%=SubHotelTypeId%>"> 
						<input type="hidden" name="ThirdHotelTypeId" value="<%=ThirdHotelTypeId%>">
						<input type="hidden" name="allgroup" value="<%=allgroup%>">
						<input type="hidden" name="allsubgroup" value="<%=allsubgroup%>">
</form>
</table></td>
                                    <td width="18"  valign="top"  background="<%=path_link%>_images/box/box-line-right.gif"><img src="<%=path_link%>_images/SPACER.GIF" width="18" height="30"></td>
                                  </tr>
                                  <tr> 
                                    <td width="18" height="12" valign="top"  background="<%=path_link%>_images/box/box-c-below-left.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td valign="top" background="<%=path_link%>_images/box/box-c-below-center.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td valign="top" background="<%=path_link%>_images/box/box-c-below-right.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                  </tr>
                                </table>

<%End Sub
Sub ShowHotelRoot(root0,root1,root2,root3)
if root0="" Then root0=0
if root1="" Then root1=0
if root2="" Then root2=0
if root3="" Then root3=0

Table0="TabHotelType"
Field0="HotelType"
Condition0=" Where HotelTypeId="&root0
Table1="TabSubHotelType"
Field1="SubHotelType"
Condition1=" Where SubHotelTypeId="&root1
Table2="TabThirdHotelType"
Field2="ThirdHotelType"
Condition2=" Where ThirdHotelTypeId="&root2
Table3="TabFourthHotelType"
Field3="FourthHotelType"
Condition3=" Where FourthHotelTypeId="&root3

url0="default.asp?page=Hotel&HotelTypeId="&root0
url1="default.asp?page=Hotel&HotelTypeId="&root0&"&SubHotelTypeId="&root1
url2="default.asp?page=Hotel&HotelTypeId="&root0&"&SubHotelTypeId="&root1&"&ThirdHotelTypeId="&root2

Dim RootName(3)
For i=0 to 3
RootName(i)=GetSingleField(eval("Table"&i),eval("Field"&i),eval("Condition"&i))



IF RootName(i)="0" then RootName(i)="" Else RootName(i)="<a href='"&eval("url"&i)&"' <img src='_images/arrow.gif'> "&RootName(i)&"</a> <img src='_images/arrow.gif'> "

If i=0 then DisplayRoot="<A Href='default.asp?page=home' class='textdefault'>Hotel</font></a> <img src='_images/arrow.gif'> "

DisplayRoot=DisplayRoot&RootName(i)
next
'DisplayRoot=DisplayRoot
DisplayRoot=left(DisplayRoot,len(DisplayRoot)-30)%>
<table width="100%" border="0" cellspacing="0" cellpadding="0" >
  <tr>
    <td class="textdefault"><b><font color=FFFFFF><%=DisplayRoot%></font></b></td>
  </tr>
</table>

<%End Sub
Function ShowHotelType(root0,root1,root2,root3)
if root0="" or root0="0" Then root0=0 Else numloop=0
if root1="" or root1="0" Then root1=0 Else numloop=1
if root2="" or root2="0" Then root2=0 Else numloop=2
if root3="" or root3="0" Then root3=0 Else numloop=3

Table0="TabHotelType"
Field0="HotelType"
Condition0=" Where HotelTypeId="&root0
Table1="TabSubHotelType"
Field1="SubHotelType"
Condition1=" Where SubHotelTypeId="&root1
Table2="TabThirdHotelType"
Field2="ThirdHotelType"
Condition2=" Where ThirdHotelTypeId="&root2
Table3="TabFourthHotelType"
Field3="FourthHotelType"
Condition3=" Where FourthHotelTypeId="&root3

url0="default.asp?page=Hotel&HotelTypeId="&root0
url1="default.asp?page=Hotel&HotelTypeId="&root0&"&SubHotelTypeId="&root1
url2="default.asp?page=Hotel&HotelTypeId="&root0&"&SubHotelTypeId="&root1&"&ThirdHotelTypeId="&root2

Dim RootName(3)
For i=0 to numloop
RootName(i)=GetSingleField(eval("Table"&i),eval("Field"&i),eval("Condition"&i))



IF RootName(i)<>"0" then  RootName(i)=RootName(i)


ShowHotelType=RootName(i)
next%>
<%End Function

Sub FormHotelAtoZ
HotelId=request("HotelId")
HotelTypeId=request("HotelTypeId")
SubHotelTypeId=request("SubHotelTypeId")
ThirdHotelTypeId=request("ThirdHotelTypeId")
ConditionAtoZ=" Where 1"
IF HotelTypeId<>"" Then ConditionAtoZ=ConditionAtoZ&" And HotelTypeId="&HotelTypeId
IF SubHotelTypeId<>"" Then ConditionAtoZ=ConditionAtoZ&" And SubHotelTypeId="&SubHotelTypeId
IF ThirdHotelTypeId<>"" Then ConditionAtoZ=ConditionAtoZ&" And ThirdHotelTypeId="&ThirdHotelTypeId%>
<script language="JavaScript">
function chkval()
{

	f=document.formAtoZ
	if (f.HotelId.value!='')
	{f.submit()}
}
</script>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                  <tr> 
                                    <td  background="<%=path_link%>_images/box/box-c-left.gif" width="18" height="12"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td   background="<%=path_link%>_images/box/box-c-center.gif"  height="12"><img src="<%=path_link%>_images/SPACER.GIF"  height="12"></td>
                                    <td   background="<%=path_link%>_images/box/box-c-right.gif" width="18" height="12"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                  </tr>
                                  <tr> 
                                    
    <td width="18"  valign="top" background="<%=path_link%>_images/box/box-line-left.gif"><img src="<%=path_link%>_images/SPACER.GIF" width="18" height="30"></td>
                                    
                
    <td valign="top" width="100%" bgcolor="#FFFFFF">
<table width="100%" border="0" cellpadding="0" cellspacing="0" class="text1" align="center">
        <form name=formAtoZ action="default.asp">
          <tr> 
            <td width="120">Hotel Name A-Z</td>
            <td> <%
		Table2= "TabHotelRent"
		order=" Order by HotelName"
		Condition2= ConditionAtoZ&order
		SelectName= "HotelId"
		FieldValue= "HotelId"
		FieldDesc= "HotelName"
		ValueExp=HotelId
		CJava= "onchange='chkval()'"
		Str_Property= ""
		 Call ListBox("=========== Hotel Name A-Z ===========",Table2,Condition2,SelectName,FieldValue,FieldDesc,ValueExp,CJava,Str_Property)
%> </td>
          </tr>
		  <input type="hidden" name="page" value="HotelInfo">
		<input type="hidden" name="HotelTypeId" value="<%=HotelTypeId%>"> 
		<input type="hidden" name="SubHotelTypeId" value="<%=SubHotelTypeId%>"> 
		<input type="hidden" name="ThirdHotelTypeId" value="<%=ThirdHotelTypeId%>">
        </form>
      </table></td>
                                    <td width="18"  valign="top"  background="<%=path_link%>_images/box/box-line-right.gif"><img src="<%=path_link%>_images/SPACER.GIF" width="18" height="30"></td>
                                  </tr>
                                  <tr> 
                                    <td width="18" height="12" valign="top"  background="<%=path_link%>_images/box/box-c-below-left.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td valign="top" background="<%=path_link%>_images/box/box-c-below-center.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                    <td valign="top" background="<%=path_link%>_images/box/box-c-below-right.gif"><img src="<%=path_link%>_images/spacer.gif" width="10" height="12"></td>
                                  </tr>
                                </table>		
<%End Sub



Function Gen_MenuPic(arr_topic,arr_link,menuid)
%>
<table width="72%" border="0" cellspacing="0" cellpadding="0" STYLE="filter:progid:DXImageTransform.Microsoft.gradient(GradientType=0,startColorstr='#FFFFFF', endColorstr='#2B9CF9')">
        <%For i=0 to Ubound(arr_topic)%>
		<tr> 
          <td>&nbsp;&nbsp;
		  <%if Cint(i)=Cint(menuid) Then%>
		  <a href="<%=arr_link(i)%>&menuid=<%=i%>" onMouseOver="MM_swapImage('Image<%=i+1%>','','_images/menu/over/<%=arr_topic(i)%>',1)" onMouseOut="MM_swapImgRestore()"><img src="<%=path_link%>_images/menu/over/<%=arr_topic(i)%>" border=0 id="Image<%=i+1%>"></a>
		  <%Else%>
		  <a href="<%=arr_link(i)%>&menuid=<%=i%>" onMouseOver="MM_swapImage('Image<%=i+1%>','','_images/menu/over/<%=arr_topic(i)%>',1)" onMouseOut="MM_swapImgRestore()"><img src="<%=path_link%>_images/menu/<%=arr_topic(i)%>" border=0 id="Image<%=i+1%>"></a>
		  <%End IF%>
		  </td>
        </tr>
		<%next%>
          <td><img src="_images/menu_bottom_03.gif"></td>
        </tr>
      </table>
	  
<%
		End Function
		
Function MenuBottomLogin(Login,NewUser,ForgotPassword,ChangePassword) 'Exam:MenuBottomLogin(1,1,0)
%>
<table width="400" border="0" cellpadding="0" cellspacing="0" class="fontmenu" align="center">
    <tr> 
      <td><img src="_images/key.gif"  align="absmiddle">
	  <% if Login=1 Then%>
	  <a href="default.asp?page=login">เข้าสู่ระบบ</a>
	  <%Else%>
	  <b>เข้าสู่ระบบ</b>
	  <%End if%>
	  </td>
      <td><img src="_images/newuser.gif"  align="absmiddle">
	  <% if NewUser=1 Then%>
	  <a href="default.asp?page=register">สมัครสมาชิก</a>
	  <%Else%>
	  <b>สมัครสมาชิก</b>
	  <%End if%>
	  </td>
      <td><img src="_images/forgot.gif"  align="absmiddle"> 
        <% if ForgotPassword=1 Then%>
		<a href="default.asp?page=forgotpassword">ลืมรหัสผ่าน</a>
	  <%Else%>
	  <b>ลืมรหัสผ่าน</b>
	  <%End if%>
		</td>
      <td><img src="_images/change.gif"  align="absmiddle"> 
        <% if ChangePassword=1 Then%>
		<a href="default.asp?page=changepassword">เปลี่ยนรหัสผ่าน</a>
	  <%Else%>
	  <b>เปลี่ยนรหัสผ่าน</b>
	  <%End if%>
		</td>
    </tr>
  </table>
  <%End Function%>


<%Sub InsertIink(RecordsetName,LinkName,EndLinkName,FieldIdName)

if RecordsetName("Active")=0 and session("member")="" then lockimg="<img src="&path_link&"_images/template/Untitled-1.gif borcer=0 alt='กรุณา login เข้าสู่ระบบก่อน' align=absmiddle>"


	if left(RecordsetName("Desc"),7)="http://" then 
	LinkName="<a href=javascript:checklogin('"&encode(RecordsetName("Desc"))&"','"&session("member")&"',"&RecordsetName("Active")&")>"
	EndLinkName="</a> "&lockimg
	elseif RecordsetName("Desc")<>"" then 
	LinkName="<a href=javascript:checklogin('"&encode("default.asp?page=data_detail&"&FieldIdName&"="&RecordsetName(FieldIdName))&"','"&session("member")&"',"&RecordsetName("Active")&")>"
	EndLinkName="</a> "&lockimg
	elseif RecordsetName("Img")<>"" then 
	LinkName="<a  href=javascript:checklogin('"&encode(path_link&Pathallpicture_front&RecordsetName("Img"))&"','"&session("member")&"',"&RecordsetName("Active")&",'_blank')>"
	EndLinkName="</a> "&lockimg
	else
    LinkName=""
	EndLinkName=""
	End if
End sub

Sub InsertIinkEx(RecordsetName,LinkName,EndLinkName,FieldIdName)

if RecordsetName("Active")=0 and session("member")="" then lockimg="<img src="&path_link&"_images/template/Untitled-1.gif borcer=0 alt='กรุณา login เข้าสู่ระบบก่อน' align=absmiddle>"


	if left(RecordsetName("Desc"),7)="http://" then 
	LinkName="<a href=javascript:checklogin('"&encode(RecordsetName("Desc"))&"','"&session("member")&"',"&RecordsetName("Active")&"),'_blank'>"
	EndLinkName="</a> "&lockimg
	elseif RecordsetName("Desc")<>"" then 
	LinkName="<a href=javascript:checklogin('"&encode("default.asp?page=data_detail&"&FieldIdName&"="&RecordsetName(FieldIdName))&"','"&session("member")&"',"&RecordsetName("Active")&",'_blank')>"
	EndLinkName="</a> "&lockimg
	elseif RecordsetName("Img")<>"" then 
	LinkName="<a  href=javascript:checklogin('"&encode(path_link&Pathallpicture_front&RecordsetName("Img"))&"','"&session("member")&"',"&RecordsetName("Active")&",'_blank')>"
	EndLinkName="</a> "&lockimg
	else
    LinkName=""
	EndLinkName=""
	End if
End sub

sub notelock%>
      
<div align="center"><span class="textbox"><font color="#0066FF">สัญลักษณ์ <img src="_images/template/Untitled-2.gif" align="absmiddle"  width="15"> 
  หมายถึง ต้อง Login เข้าสู่ระบบก่อนจึงจะสามารถเปิดอ่านได้</font></span></div>
<font color="#0066FF"> 
<%end sub%>
<%website=":::K M::: Knowledge Management"%>
</font>