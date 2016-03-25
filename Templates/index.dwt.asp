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

Repeat1__numRows = 6
Repeat1__index = 0
sach_numRows = sach_numRows + Repeat1__numRows
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!-- TemplateBeginEditable name="doctitle" -->
<title>Free Leoshop Website Template | Home :: w3layouts</title>
<!-- TemplateEndEditable -->
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link href="../guest/css/mystyle.css" rel="stylesheet" type="text/css" />
<link href="../guest/css/style.css" rel="stylesheet" type="text/css" media="all" />
<link href="../guest/css/form.css" rel="stylesheet" type="text/css" media="all" />
<link href='http://fonts.googleapis.com/css?family=Exo+2' rel='stylesheet' type='text/css'>
<script type="text/javascript" src="../guest/js/jquery1.min.js"></script>
<!-- start menu -->
<link href="../guest/css/megamenu.css" rel="stylesheet" type="text/css" media="all" />
<script type="text/javascript" src="../guest/js/megamenu.js"></script>
<script>$(document).ready(function(){$(".megamenu").megamenu();});</script>
<!--start slider -->
    <link rel="stylesheet" href="../guest/css/fwslider.css" media="all">
    <script src="../guest/js/jquery-ui.min.js"></script>
    <script src="../guest/js/css3-mediaqueries.js"></script>
    <script src="../guest/js/fwslider.js"></script>
<!--end slider -->
<script src="../guest/js/jquery.easydropdown.js"></script>
<!-- TemplateBeginEditable name="head" -->
<!-- TemplateEndEditable -->
</head>
<body>
     <div class="header-top">
	   <div class="wrap"> 
			  
			 <div class="cssmenu">
				<ul>
					<li class="active"><a href="../guest/DangXuat user.asp">Đăng xuất</a></li> |
					<li><a href="../guest/DangNhap.asp">Đăng Nhập</a></li> |
					<li><a href="../guest/Dangky.asp">Đăng Ký</a></li>
				</ul>
			</div>
			<div class="clear"></div>
	   </div>
	</div>
	<div class="header-bottom">
	    <div class="wrap">
			<div class="header-bottom-left">
				<div class="logo">
					<a href="../guest/index.html"><img src="../guest/images/logo.png" alt=""/></a>
				</div>
				<div class="menu">
	            <ul class="megamenu skyblue">
			<li class="active grid"><a href="../guest/index.html">TRANG CHỦ</a></li>
			<li><a class="color4" href="#">THỂ LOẠI</a>
			  <div class="megapanel">
					<div class="row">
						<div class="col1">
							<div class="h_nav">
								<h4>Contact Lenses</h4>
								<ul>
									<li><a href="../guest/womens.html">Daily-wear soft lenses</a></li>
									<li><a href="../guest/womens.html">Extended-wear</a></li>
									<li><a href="../guest/womens.html">Lorem ipsum </a></li>
									<li><a href="../guest/womens.html">Planned replacement</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Sun Glasses</h4>
								<ul>
									<li><a href="../guest/womens.html">Heart-Shaped</a></li>
									<li><a href="../guest/womens.html">Square-Shaped</a></li>
									<li><a href="../guest/womens.html">Round-Shaped</a></li>
									<li><a href="../guest/womens.html">Oval-Shaped</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Eye Glasses</h4>
								<ul>
									<li><a href="../guest/womens.html">Anti Reflective</a></li>
									<li><a href="../guest/womens.html">Aspheric</a></li>
									<li><a href="../guest/womens.html">Bifocal</a></li>
									<li><a href="../guest/womens.html">Hi-index</a></li>
									<li><a href="../guest/womens.html">Progressive</a></li>
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
									<li><a href="../guest/mens.html">Daily-wear soft lenses</a></li>
									<li><a href="../guest/mens.html">Extended-wear</a></li>
									<li><a href="../guest/mens.html">Lorem ipsum </a></li>
									<li><a href="../guest/mens.html">Planned replacement</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Sun Glasses</h4>
								<ul>
									<li><a href="../guest/mens.html">Heart-Shaped</a></li>
									<li><a href="../guest/mens.html">Square-Shaped</a></li>
									<li><a href="../guest/mens.html">Round-Shaped</a></li>
									<li><a href="../guest/mens.html">Oval-Shaped</a></li>
								</ul>	
							</div>							
						</div>
						<div class="col1">
							<div class="h_nav">
								<h4>Eye Glasses</h4>
								<ul>
									<li><a href="../guest/mens.html">Anti Reflective</a></li>
									<li><a href="../guest/mens.html">Aspheric</a></li>
									<li><a href="../guest/mens.html">Bifocal</a></li>
									<li><a href="../guest/mens.html">Hi-index</a></li>
									<li><a href="../guest/mens.html">Progressive</a></li>
								</ul>	
							</div>												
						</div>
					</div>
				</li>
			</ul>
			</div>
		</div>
	   <div class="header-bottom-right">
         <div class="search">	  
				<input type="text" name="s" class="textbox" value="Search" onfocus="this.value = '';" onblur="if (this.value == '') {this.value = 'Search';}">
				<input type="submit" value="Subscribe" id="submit" name="submit">
				<div id="response"> </div>
		 </div>
	  <div class="tag-list">
		<ul class="icon1 sub-icon1 profile_img">
			<li><a class="active-icon c2" href="#"> </a>
				<ul class="sub-icon1 list">
					<li><h3>No Products</h3><a href=""></a></li>
					<li><p>Lorem ipsum dolor sit amet, consectetuer  <a href="">adipiscing elit, sed diam</a></p></li>
				</ul>
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
                    <img src="../guest/images/banner.jpg" alt=""/>
                <!-- /Slide image -->
            </div>
            <!-- /Duplicate to create more slides -->
            <div class="slide">
                <img src="../guest/images/banner1.jpg" alt=""/>
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
		  	<h2 class="head">Đặc sắc		    </h2>
		  	<div class="section group">
			  <!-- TemplateBeginEditable name="nội dung" -->
              <p>
                <% 
While ((Repeat1__numRows <> 0) AND (NOT sach.EOF)) 
%>
                <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  sach.MoveNext()
Wend
%>
              </p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp;</p>
              <p>&nbsp; </p>
              <!-- TemplateEndEditable -->
			  <div class="clear"></div>
			</div>			 						 			    
		  </div>
			<div class="rsidebar span_1_of_left">
				<div class="top-border"> </div>
				 <div class="border">
	             <link href="../guest/css/default.css" rel="stylesheet" type="text/css" media="all" />
	             <link href="../guest/css/nivo-slider.css" rel="stylesheet" type="text/css" media="all" />
				  <script src="../guest/js/jquery.nivo.slider.js"></script>
				    <script type="text/javascript">
				    $(window).load(function() {
				        $('#slider').nivoSlider();
				    });
				    </script>
		    <div class="slider-wrapper theme-default">
              <div id="slider" class="nivoSlider">
                <img src="../guest/images/t-img1.jpg"  alt="" />
               	<img src="../guest/images/t-img2.jpg"  alt="" />
                <img src="../guest/images/t-img3.jpg"  alt="" />
              </div>
             </div>
              <div class="btn"><a href="../guest/single.html">Check it Out</a></div>
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
					  <li><img src="../guest/images/2.png"><span class="f-text">Miễn Phí Giao Hàng</span><div class="clear"></div></li>
					</ul>
				</div>
				<div class="col_1_of_2 span_1_of_2">
					<ul class="f-list">
					  <li><img src="../guest/images/3.png"><span class="f-text">Điện Thoại 0908070605 </span><div class="clear"></div></li>
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
