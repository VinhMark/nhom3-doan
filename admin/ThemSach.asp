<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/connect.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="XemSach.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
'*********************************
'*                               *
'*  INSERT RECORD AND UPLOAD     *
'*  http://www.dwzone.it         *
'*                               *
'*********************************
	server.ScriptTimeout = 5400

	Dim RG_altVal, RG_columns, RG_Cong, RG_dbValues, RG_dbValuesTmp, RG_delim, RG_editCmd, RG_editQuery, RG_editQueryTmp, RG_emptyVal, RG_Ext, RG_Extensions, RG_fields, RG_FieldValueTmp, RG_FileDel, RG_FileExt, RG_formVal, RG_FS, RG_i, RG_L, RG_Len, RG_Max, RG_Name, RG_New, RG_newName, RG_Num, RG_Path, RG_Rec, RG_ret, RG_Save, RG_tableValues, RG_tableValuesTmp, RG_tst, RG_typeArray, RG_z, UploadStatus, NumFile
	Dim RG_Connection, RG_editColumn, RG_recordId, Form, editAction, editRedirectUrl, RG_Files, RG_formName, UploadType, ParamVal, ParamList, MaxFieldNumber, TmpVal, x, y, Key, ProgressBar
	Dim tmpField_Name(), tmpValue_Name(), tmpField_Size(), tmpValue_Size(), tmpField_Thumb(), tmpValue_Thumb(), QtyRecord, TotalFileSize, valueToRedirectSend
	
	Set Form = New ASPForm
	Dim UploadID
	UploadID = Form.NewUploadID
	ProgressBar = "progress-std.asp"
	TotalFileSize = ""
  	editRedirectUrl = "XemSach.asp"
	RG_Connection = MM_connect_STRING
	RG_editTable = "dbo.Sach"	
	RG_Files = "/img;1;;;;0;HinhAnh;1;;;;0;;;;;;fimg@_@_@1@_@_@ @_@_@@_@_@../"
	RG_formName = "form1"
	UploadType="Insert"
	UploadStatus = ""
	valueToRedirectSend = ""
	NumFile = 0

	if len(Request.QueryString("UploadID"))>0 then
		Form.UploadID = Request.QueryString("UploadID")
	end if
	
	if (Request.QueryString <> "") Then
	 	editAction = CStr(Request.ServerVariables("SCRIPT_NAME")) & "?" & Request.QueryString & "&UploadID=" & UploadID
	else
	 	editAction = CStr(Request.ServerVariables("SCRIPT_NAME")) & "?UploadID=" & UploadID	
	End If
	
Const fsCompletted  = 0

If Form.State = fsCompletted Then 
  if Form.State = 0 then

		Set ParamVal = CreateObject("Scripting.Dictionary")
		tmp = split(RG_Files,"@_@_@")
		ParamList = split(tmp(0),"|")
		MaxFieldNumber = ubound(ParamList)
		for x=0 to Ubound(ParamList)
			TmpVal = Split(ParamList(x),";")
			for y=0 to ubound(TmpVal)
				Key = right("00" & cstr(x),3) & cstr(y)
				ParamVal.add Key, TmpVal(y)
			next
		next
		Form.Files.Save

		RG_fieldsStr  = "txtTenSach|value|txtTinhTrang|value|txtHienthi|value|btnTacGia|value|btnTheLoai|value|txtNXB|value|txtGia|value|txtSoLuong|value|txtMoTa|value"
  		RG_columnsStr = "TenSach|',none,''|TinhTrang|none,none,NULL|HienThi|none,none,NULL|TacGia|none,none,NULL|TheLoai|none,none,NULL|NXB|',none,''|Gia|none,none,NULL|SoLuong|none,none,NULL|MoTa|',none,''"
		Form.Files.DataBaseInsert

		response.write(getRedirect())
		response.end
  End If
ElseIf Form.State > 10 then
  response.write ""
End If


function GetFolderName(str):  GetFolderName = Ris : end function

function myGetFileName(str):  myGetFileName = Ris : end function
%>
<%
Dim tacgia
Dim tacgia_cmd
Dim tacgia_numRows

Set tacgia_cmd = Server.CreateObject ("ADODB.Command")
tacgia_cmd.ActiveConnection = MM_connect_STRING
tacgia_cmd.CommandText = "SELECT * FROM dbo.TacGia" 
tacgia_cmd.Prepared = true

Set tacgia = tacgia_cmd.Execute
tacgia_numRows = 0
%>
<%
Dim theloai
Dim theloai_cmd
Dim theloai_numRows

Set theloai_cmd = Server.CreateObject ("ADODB.Command")
theloai_cmd.ActiveConnection = MM_connect_STRING
theloai_cmd.CommandText = "SELECT * FROM dbo.TheLoai" 
theloai_cmd.Prepared = true

Set theloai = theloai_cmd.Execute
theloai_numRows = 0
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"><!-- InstanceBegin template="/Templates/template-admin.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta name="description" content="Creative - Bootstrap 3 Responsive Admin Template">
    <meta name="author" content="GeeksLabs">
    <meta name="keyword" content="Creative, Dashboard, Admin, Template, Theme, Bootstrap, Responsive, Retina, Minimal">
    <link rel="shortcut icon" href="img/favicon.png">

    <title>Admin LeoShop</title>

    <!-- Bootstrap CSS -->  
    <link href="../css/stylecuatoi.css" rel="stylesheet" type="text/css" />  
    <link href="../css/bootstrap.min.css" rel="stylesheet">
    <!-- bootstrap theme -->
    <link href="../css/bootstrap-theme.css" rel="stylesheet">
    <!--external css-->
    <!-- font icon -->
    <link href="../css/elegant-icons-style.css" rel="stylesheet" />
    <link href="../css/font-awesome.min.css" rel="stylesheet" />    
    <!-- full calendar css-->
    <link href="../assets/fullcalendar/fullcalendar/bootstrap-fullcalendar.css" rel="stylesheet" />
	<link href="../assets/fullcalendar/fullcalendar/fullcalendar.css" rel="stylesheet" />
    <!-- easy pie chart-->
    <link href="../assets/jquery-easy-pie-chart/jquery.easy-pie-chart.css" rel="stylesheet" type="text/css" media="screen"/>
    <!-- owl carousel -->
    <link rel="stylesheet" href="../css/owl.carousel.css" type="text/css">
	<link href="../css/jquery-jvectormap-1.2.2.css" rel="stylesheet">
    <!-- Custom styles -->
	<link rel="stylesheet" href="../css/fullcalendar.css">
	<link href="../css/widgets.css" rel="stylesheet">
    <link href="../css/style.css" rel="stylesheet">
    <link href="../css/style-responsive.css" rel="stylesheet" />
	<link href="../css/xcharts.min.css" rel=" stylesheet">	
	<link href="../css/jquery-ui-1.10.4.min.css" rel="stylesheet">
    <!-- HTML5 shim and Respond.js IE8 support of HTML5 -->
    <!--[if lt IE 9]>
      <script src="../js/html5shiv.js"></script>
      <script src="../js/respond.min.js"></script>
      <script src="../js/lte-ie7.js"></script>
    <![endif]-->
  </head>

<body>
  <!-- container section start -->
  <section id="container" class="">
     
      
      <header class="header dark-bg">
            <div class="toggle-nav">
                <div class="icon-reorder tooltips" data-original-title="Toggle Navigation" data-placement="bottom"><i class="icon_menu"></i></div>
            </div>

            <!--logo start-->
            <a href="index.asp" class="logo"><span class="lite">Admin</span></a>
            <!--logo end-->

            <div class="nav search-row" id="top_menu">
                <!--  search form start -->
                <ul class="nav top-menu">                    
                    <li>
                        <!--<form class="navbar-form">
                            <input class="form-control" placeholder="Search" type="text">
                        </form>-->
                    </li>                    
                </ul>
                <!--  search form end -->                
            </div>

            <div class="top-nav notification-row">                
                <!-- notificatoin dropdown start-->
                <ul class="nav pull-right top-menu">
                    
                    <!-- task notificatoin start -->
                    <li id="task_notificatoin_bar" class="dropdown">
                       
                        <ul class="dropdown-menu extended tasks-bar">
                            <div class="notify-arrow notify-arrow-blue"></div>
                            <li>
                              
                            </li>
                            <li>
                                    <div class="progress progress-striped">
                                        
                                    </div>
                                </a>
                            </li>
                            <li>
                                <a href="#">
                                    <div class="task-info">
                                        
                                    </div>
                                    <div class="progress progress-striped">
                                        
                                    </div>
                                </a>
                            </li>
                            <li>
                                <a href="#">
                                    
                                    <div class="progress progress-striped">
                                        
                                    </div>
                                </a>
                            </li>
                            <li>
                                <a href="#">
                                    <div class="task-info">
                                      
                                    </div>
                                    <div class="progress progress-striped">
                                        
                                    </div>
                                </a>
                            </li>
                            <li>
                                <a href="#">
                                  
                                    <div class="progress progress-striped active">
                                      
                                    </div>

                                </a>
                            </li>
                            
                        </ul>
                    </li>
                    <!-- task notificatoin end -->
                    <!-- inbox notificatoin start-->
                    <li id="mail_notificatoin_bar" class="dropdown">
                       
                        <ul class="dropdown-menu extended inbox">
                            <div class="notify-arrow notify-arrow-blue"></div>
                            
                            <li>
                                
                               					
                            </li>
                            <li>
                               
                            </li>
                            <li>
                                
                            </li>
                            <li>
                                
                            </li>
                            <li>
                               
                            </li>
                        </ul>
                    </li>
                    <!-- inbox notificatoin end -->
                    <!-- alert notification start-->
                    <li id="alert_notificatoin_bar" class="dropdown">
                        
                        
                    </li>
                    <!-- alert notification end-->
                    <!-- user login dropdown start-->
                    <li class="dropdown">
                        
                            
                            <span class="username"><a href="dangnhap.asp">Đăng xuất</a></span>
                        </a>
                        
                    </li>
                    <!-- user login dropdown end -->
                </ul>
                <!-- notificatoin dropdown end-->
            </div>
      </header>      
      <!--header end-->

      <!--sidebar start-->
      <aside>
          <div id="sidebar"  class="nav-collapse ">
              <!-- sidebar menu start-->
              <ul class="sidebar-menu">                
                  <li class="active">
                      <a class="" href="index.asp">
                          <i class="icon_house_alt"></i>
                          <span>Trang chủ</span></a>
                  </li>
				  <li class="sub-menu">
                      <a href="javascript:;" class="">
                          <i class="icon_document_alt"></i>
                          <span>Sản Phẩm</span>
                          <span class="menu-arrow arrow_carrot-right"></span>
                      </a>
                      <ul class="sub">
                      	  <li><a class="" href="XemSach.asp">Sách</a></li> 
                          <li><a class="" href="ThemSach.asp">Thêm Sách</a></li>                          
                          <li><a class="" href="ThemTL.asp">Thêm Thể Loại</a></li>
                          <li><a class="" href="ThemTacGia.asp">Thêm Thể Loại</a></li>
                      </ul>
                  </li>       
                  <li class="sub-menu">
                      <a href="javascript:;" class="">
                          <i class="icon_desktop"></i>
                          <span>Tài khoản</span>
                          <span class="menu-arrow arrow_carrot-right"></span>
                      </a>
                      <ul class="sub">
                      	  <li><a class="" href="TaiKhoan.asp">Khách hàng</a></li>
                          <li><a class="" href="TaiKhoanAdmin.asp">Admin</a></li> 
                      </ul>
                      
                  </li>
                  <li></li>
                  <li></li>
                             
                  <li class="sub-menu">
                      <a href="javascript:;" class="">
                          <i class="icon_table"></i>Đơn Hàng<span class="menu-arrow arrow_carrot-right"></span>
                      </a>
                      <ul class="sub">
                          <li><a class="" href="DonDatHang.asp">Đơn Đặt Hàng</a></li>
                          <li><a class="" href="CTDDH.asp">CTDDH</a></li>
                      </ul>
                      
                          
                     

                      <ul class="sub">
                          <li><a class="" href="CTDDH.asp">CTDDH</a></li>
                      </ul>
                  </li>
                  
                  <li class="sub-menu">
                    <ul class="sub">                          
                      <li><a class="" href="profile.html">Profile</a></li>
                          <li><a class="" href="login.html"><span>Login Page</span></a></li>
                          <li><a class="" href="blank.html">Blank Page</a></li>
                          <li><a class="" href="404.html">404 Error</a></li>
                      </ul>
                  </li>
                  
              </ul>
              <!-- sidebar menu end-->
          </div>
      </aside>
      <!--sidebar end-->
      
      <!--main content start-->
      <section id="main-content">
          <section class="wrapper">            
              <!--overview start-->
			  <div class="row"><!-- InstanceBeginEditable name="nội dung" -->
              <div class="noidung">
			    <table class="theten" width="98%" border="0" cellspacing="0" cellpadding="0">
			      <tr>
			        <td height="44"><strong><h2>Thêm Sách</h2></strong></td>
		          </tr>
                  </table>
			      <tr>
			        <td width="100%" height="44"><form onsubmit="return ProgressBar()" action="<%=editAction%>" method="post" enctype="multipart/form-data" name="form1" id="form1">
			          <table class="thethongtin" width="98%" border="0" cellspacing="0" cellpadding="0">
			            <tr>
			              <td width="7%">&nbsp;</td>
			              <td width="29%"><h2>Tên sách</h2></td>
			              <td width="64%"><label for="txtTenSach"></label>
                          <div>
		                  	<input type="text" name="txtTenSach" id="txtTenSach" />
                          </div>
		                  <input name="txtTinhTrang" type="hidden" id="txtTinhTrang" value="1" />
		                  <input name="txtHienthi" type="hidden" id="txtHienthi" value="1" /></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td><h3>Tác giả</h3></td>
			              <td><label for="btnTacGia"></label>
			                <div><select name="btnTacGia" id="btnTacGia">
			                  <%
While (NOT tacgia.EOF)
%>
			                  <option value="<%=(tacgia.Fields.Item("MaTG").Value)%>"><%=(tacgia.Fields.Item("TenTG").Value)%></option>
			                  <%
  tacgia.MoveNext()
Wend
If (tacgia.CursorType > 0) Then
  tacgia.MoveFirst
Else
  tacgia.Requery
End If
%>
                          </select></div></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td><h3>Thể loại</h3></td>
			              <td><label for="btnTheLoai"></label>
			                <select name="btnTheLoai" id="btnTheLoai">
			                  <%
While (NOT theloai.EOF)
%>
			                  <option value="<%=(theloai.Fields.Item("MaTL").Value)%>"><%=(theloai.Fields.Item("TenTL").Value)%></option>
			                  <%
  theloai.MoveNext()
Wend
If (theloai.CursorType > 0) Then
  theloai.MoveFirst
Else
  theloai.Requery
End If
%>
                          </select></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td><h3>NXB</h3></td>
			              <td><label for="txtNXB"></label>
		                  <input type="text" name="txtNXB" id="txtNXB" /></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td><h3>Hình ảnh</h3></td>
			              <td><label for="fimg"></label>
		                  <input type="file" name="fimg" id="fimg" /></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td><h3>Giá</h3></td>
			              <td><label for="txtGia"></label>
		                  <input type="text" name="txtGia" id="txtGia" /></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td><h3>Số lượng<h3></td>
			              <td><label for="txtSoLuong"></label>
		                  <input type="text" name="txtSoLuong" id="txtSoLuong" /></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td><h3>Mô tả</h3></td>
			              <td><label for="txtMoTa"></label>
		                  <input type="text" name="txtMoTa" id="txtMoTa" /></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td>&nbsp;</td>
			              <td>&nbsp;</td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td><input type="submit" name="Submit" id="button" value="Thêm" class="btn"/></td>
			              <td><input type="reset" name="button2" id="button2" value="Hủy" class="btn"/></td>
		                </tr>
			            <tr>
			              <td>&nbsp;</td>
			              <td>&nbsp;</td>
			              <td>&nbsp;</td>
		                </tr>
		              </table>
		            </form></td>
		          </tr>
		      
              </div>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
			    <p>&nbsp;</p>
		    <!-- InstanceEndEditable --></div>
              
            <div class="row"><!--/.col--><!--/.col--><!--/.col--><!--/.col-->
				
			</div><!--/.row-->
		
			
           <div class="row"></div>  
            
		  
		  <!-- Today status end -->
			
              
				
			<div class="row"><!--/col--><!--/col--></div>

                    
                   
                <!-- statics end -->
              
            
				

              <!-- project team & activity start -->
          <div class="row"></div>
          <p>&nbsp;	</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p>&nbsp;</p>
          <p><br>
            <br>
            
          </p>
          <div class="row"></div> 
              <!-- project team & activity end -->

          </section>
      </section>
      <!--main content end-->
  </section>
  <!-- container section start -->

    <!-- javascripts -->
    <script src="../js/jquery.js"></script>
	<script src="../js/jquery-ui-1.10.4.min.js"></script>
    <script src="../js/jquery-1.8.3.min.js"></script>
    <script type="text/javascript" src="../js/jquery-ui-1.9.2.custom.min.js"></script>
    <!-- bootstrap -->
    <script src="../js/bootstrap.min.js"></script>
    <!-- nice scroll -->
    <script src="../js/jquery.scrollTo.min.js"></script>
    <script src="../js/jquery.nicescroll.js" type="text/javascript"></script>
    <!-- charts scripts -->
    <script src="../assets/jquery-knob/js/jquery.knob.js"></script>
    <script src="../js/jquery.sparkline.js" type="text/javascript"></script>
    <script src="../assets/jquery-easy-pie-chart/jquery.easy-pie-chart.js"></script>
    <script src="../js/owl.carousel.js" ></script>
    <!-- jQuery full calendar -->
    <<script src="../js/fullcalendar.min.js"></script> <!-- Full Google Calendar - Calendar -->
	<script src="../assets/fullcalendar/fullcalendar/fullcalendar.js"></script>
    <!--script for this page only-->
    <script src="../js/calendar-custom.js"></script>
	<script src="../js/jquery.rateit.min.js"></script>
    <!-- custom select -->
    <script src="../js/jquery.customSelect.min.js" ></script>
	<script src="../assets/chart-master/Chart.js"></script>
   
    <!--custome script for all page-->
    <script src="../js/scripts.js"></script>
    <!-- custom script for this page-->
    <script src="../js/sparkline-chart.js"></script>
    <script src="../js/easy-pie-chart.js"></script>
	<script src="../js/jquery-jvectormap-1.2.2.min.js"></script>
	<script src="../js/jquery-jvectormap-world-mill-en.js"></script>
	<script src="../js/xcharts.min.js"></script>
	<script src="../js/jquery.autosize.min.js"></script>
	<script src="../js/jquery.placeholder.min.js"></script>
	<script src="../js/gdp-data.js"></script>	
	<script src="../js/morris.min.js"></script>
	<script src="../js/sparklines.js"></script>	
	<script src="../js/charts.js"></script>
	<script src="../js/jquery.slimscroll.min.js"></script>
  <script>

      //knob
      $(function() {
        $(".knob").knob({
          'draw' : function () { 
            $(this.i).val(this.cv + '%')
          }
        })
      });

      //carousel
      $(document).ready(function() {
          $("#owl-slider").owlCarousel({
              navigation : true,
              slideSpeed : 300,
              paginationSpeed : 400,
              singleItem : true

          });
      });

      //custom select box

      $(function(){
          $('select.styled').customSelect();
      });
	  
	  /* ---------- Map ---------- */
	$(function(){
	  $('#map').vectorMap({
	    map: 'world_mill_en',
	    series: {
	      regions: [{
	        values: gdpData,
	        scale: ['#000', '#000'],
	        normalizeFunction: 'polynomial'
	      }]
	    },
		backgroundColor: '#eef3f7',
	    onLabelShow: function(e, el, code){
	      el.html(el.html()+' (GDP - '+gdpData[code]+')');
	    }
	  });
	});



  </script>

  </body>
<!-- InstanceEnd --></html>
<!--#include file="../UploadFiles/Upload.asp" -->
<!--#include file="../UploadFiles/UploadAdvanced.asp" -->
<%
tacgia.Close()
Set tacgia = Nothing
%>
<%
theloai.Close()
Set theloai = Nothing
%>
