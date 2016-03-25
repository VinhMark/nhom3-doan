<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="../Connections/connect.asp" -->
<%
Dim sach
Dim sach_cmd
Dim sach_numRows

Set sach_cmd = Server.CreateObject ("ADODB.Command")
sach_cmd.ActiveConnection = MM_connect_STRING
sach_cmd.CommandText = "SELECT * FROM dbo.Sach WHERE TinhTrang=1 AND HienThi=1 ORDER BY MaSach DESC" 
sach_cmd.Prepared = true

Set sach = sach_cmd.Execute
sach_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = 8
Repeat1__index = 0
sach_numRows = sach_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Free Leoshop Website Template | Home :: w3layouts</title>
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="css/mystyle.css" rel="stylesheet" type="text/css" />
<link href="css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href="css/form.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Exo+2' rel='stylesheet' type='text/css'>


<script type="text/javascript" src="js/jquery1.min.js"></script>
<!-- start menu -->
<link href="css/megamenu.css" rel="stylesheet" type="text/css" media="all" />
<script type="text/javascript" src="js/megamenu.js"></script>
<script>$(document).ready(function(){$(".megamenu").megamenu();});</script>
<!--start slider -->
    <link rel="stylesheet" href="css/fwslider.css" media="all">
    <script src="js/jquery-ui.min.js"></script>
    <script src="js/css3-mediaqueries.js"></script>
    <script src="js/fwslider.js"></script>
<!--end slider -->
<script src="js/jquery.easydropdown.js"></script>
<script src="js/jquery.magnific-popup.js" type="text/javascript"></script>
<link href="css/magnific-popup.css" rel="stylesheet" type="text/css">
		<script>
			$(document).ready(function() {
				$('.popup-with-zoom-anim').magnificPopup({
					type: 'inline',
					fixedContentPos: false,
					fixedBgPos: true,
					overflowY: 'auto',
					closeBtnInside: true,
					preloader: false,
					midClick: true,
					removalDelay: 300,
					mainClass: 'my-mfp-zoom-in'
			});
		});
		</script>
</head>
<body>
     <div class="header-top">
	   <div class="wrap"> 
			  
			 <div class="cssmenu">
				<ul>
					<li><a href="DangXuat user.asp">Đăng xuất</a></li>
					 |
					<li></li>
				</ul>
			</div>
			<div class="clear"></div>
	   </div>
	</div>
	<div class="header-bottom">
	    <div class="wrap">
			<div class="header-bottom-left">
				<div class="logo">
					<a href="index.asp"><img src="images/logo.png" alt=""/></a>
				</div>
				<div class="menu">
	            <ul class="megamenu skyblue">
			<li class="active grid"><a href="index.asp">TRANG CHỦ</a></li>
			<li><a class="color4" href="#">THỂ LOẠI</a>
			  <div class="megapanel">
					<div class="row">
						<div class="col1">
							<div class="h_nav">
								<h4>Contact Lenses</h4>
								<ul>
									<li><a href="womens.html">Daily-wear soft lenses</a></li>
									<li><a href="womens.html">Extended-wear</a></li>
									<li><a href="womens.html">Lorem ipsum </a></li>
									<li><a href="womens.html">Planned replacement</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Sun Glasses</h4>
								<ul>
									<li><a href="womens.html">Heart-Shaped</a></li>
									<li><a href="womens.html">Square-Shaped</a></li>
									<li><a href="womens.html">Round-Shaped</a></li>
									<li><a href="womens.html">Oval-Shaped</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Eye Glasses</h4>
								<ul>
									<li><a href="womens.html">Anti Reflective</a></li>
									<li><a href="womens.html">Aspheric</a></li>
									<li><a href="womens.html">Bifocal</a></li>
									<li><a href="womens.html">Hi-index</a></li>
									<li><a href="womens.html">Progressive</a></li>
								</ul>	
							</div>												
						</div>
					  </div>
					</div>
				</li>				
				<li><a class="color5" href="#">Hổ Trợ</a>
				<div class="megapanel">
					<div class="col1">
							<div class="h_nav">
								<h4>Contact Lenses</h4>
								<ul>
									<li><a href="mens.html">Daily-wear soft lenses</a></li>
									<li><a href="mens.html">Extended-wear</a></li>
									<li><a href="mens.html">Lorem ipsum </a></li>
									<li><a href="mens.html">Planned replacement</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Sun Glasses</h4>
								<ul>
									<li><a href="mens.html">Heart-Shaped</a></li>
									<li><a href="mens.html">Square-Shaped</a></li>
									<li><a href="mens.html">Round-Shaped</a></li>
									<li><a href="mens.html">Oval-Shaped</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Eye Glasses</h4>
								<ul>
									<li><a href="mens.html">Anti Reflective</a></li>
									<li><a href="mens.html">Aspheric</a></li>
									<li><a href="mens.html">Bifocal</a></li>
									<li><a href="mens.html">Hi-index</a></li>
									<li><a href="mens.html">Progressive</a></li>
								</ul>	
							</div>												
						</div>
					</div>
				</li>
			</ul>
			</div>
		</div>
	   <div class="header-bottom-right">
        <form action="Timkiem.asp" method="post"> 
        	<div class="search">	  
				<input name="txtTimKiem" type="text" class="textbox" id="txtTimKiem" onfocus="this.value = '';" onblur="if (this.value == '') {this.value = 'Tìm Kiếm';}" value="Tìm kiếm">
				<input type="submit" value="Tìm Kiếm" id="submit" name="submit">
				<div id="response"> </div>
		 	</div> 
         </form>
	  <div class="tag-list">
		<ul class="icon1 sub-icon1 profile_img">
			<li><a class="active-icon c2" href="Giohang.asp"> </a>
				
			</li>
		</ul>
	  </div>
    </div>
     <div class="clear"></div>
     </div>
	</div>
  <!-- start slider -->
    <div id="fwslider">
        <div class="slider_container">
            <div class="slide"> 
                <!-- Slide image -->
                    <img src="images/banner.jpg" alt=""/>
                <!-- /Slide image -->
            </div>
            <!-- /Duplicate to create more slides -->
            <div class="slide">
                <img src="images/banner1.jpg" alt=""/>
            </div>
            <!--/slide -->
        </div>
        <div class="timers"></div>
        <div class="slidePrev"><span></span></div>
        <div class="slideNext"><span></span></div>
    </div>
    <!--/slider -->
<div class="main">
	<div class="wrap">
		<div class="section group">
		  <div class="cont span_2_of_3">
		  	<h2 class="head">Sách Mới</h2>
            <div class="top-box"><!-- TemplateBeginEditable name="noi dung" -->
            <% 
While ((Repeat1__numRows <> 0) AND (NOT sach.EOF)) 
%>

  <div class="item">
    
    <div class="hinh hinhanh">
    
    	
    	<img src="<%=(sach.Fields.Item("HinhAnh").Value)%>" alt="ThongTin.asp" name="" width="226" height="309" />  
      	<div class="sale-box"><span class="on_sale title_shop">Mới</span></div>
        <form action="Thongtin.asp" method="post">
        	<div class="thean">
            <input type="image" name="imageField" id="imageField" src="images/xemthem2.png" />
        	<input name="MaSP" type="hidden" id="MaSP" value="<%=(sach.Fields.Item("MaSach").Value)%>" />
            </div>
        </form>
   
         </div>
    
   	<div class="the-trai">
          <p id="the-trai"><%=(sach.Fields.Item("TenSach").Value)%></p>
          <div class="giasach">Giá :<%=(sach.Fields.Item("Gia").Value)%> VNĐ</div>
    </div>
        <form action="Giohang.asp" method="post">
        	<div class="the-phai">
      			<input name="MaSP" type="hidden" id="MaSP" value="<%=(sach.Fields.Item("MaSach").Value)%>" />
      			<input name="TenSP" type="hidden" id="TenSP" value="<%=(sach.Fields.Item("TenSach").Value)%>" />
      			<input name="HinHanhSP" type="hidden" id="HinHanhSP" value="<%=(sach.Fields.Item("HinhAnh").Value)%>" />
                <input name="GiaSP" type="hidden" id="GiaSP" value="<%=(sach.Fields.Item("Gia").Value)%>" />
                <input type="image" name="imageField" id="imageField" src="images/cart.png" />
        	</div>
        </form>    
            
        <div class="tay"></div>
        
  	
    
            </div>
    
    
    
    
  
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  sach.MoveNext()
Wend
%>
            <!-- TemplateEndEditable -->
              <div class="clear"></div>
            </div>
          </div>
		  <div class="clear"></div>
	</div>
	</div>
</div>
   <div class="footer">
		<div class="footer-top">
			<div class="wrap">
			  <div class="section group example">
				<div class="col_1_of_2 span_1_of_2">

					<ul class="f-list">
					  <li><img src="images/2.png"><span class="f-text">Miễn Phí Giao Hàng</span><div class="clear"></div></li>
					</ul>
				</div>
				<div class="col_1_of_2 span_1_of_2">
					<ul class="f-list">
					  <li><img src="images/3.png"><span class="f-text">Điện Thoại 0908070605 </span><div class="clear"></div></li>
					</ul>
				</div>
				<div class="clear"></div>
		      </div>
			</div>
		</div>
		
		<div class="footer-bottom">
			<div class="wrap">
	             <div class="copy">
			        <p>© 2016 Sử Dụng Template Tại <a href="http://w3layouts.com" target="_blank">w3layouts</a></p>
		         </div>
			    
	      </div>
     </div>
</div>
    
   
</body>
</html>
<%
sach.Close()
Set sach = Nothing
%>
