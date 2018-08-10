<?php
include_once("includes/loginstatus.php");
// If user is already logged in
if (isset($_SESSION["userid"])) {
  header("location: account.php?id=".$_SESSION["userid"]."&dashboard=focus");
  exit();
}
?>

<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>PPBMS - Login</title>

<link href="css/bootstrap.min.css" rel="stylesheet">
<link href="css/datepicker3.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<link href="css/custom/login.css?refresh=true" rel="stylesheet">
<link rel="stylesheet" href="css/font-awesome.min.css">
<!-- FavIcon -->
<link rel="icon" type="image/png" href="img/fav.png" />
<!-- Google Fonts -->
<link href="https://fonts.googleapis.com/css?family=Quicksand" rel="stylesheet">

<!--[if lt IE 9]>
<script src="js/html5shiv.js"></script>
<script src="js/respond.min.js"></script>
<![endif]-->

</head>

<body>
	
	<div class="row">
		<div class="col-xs-10 col-xs-offset-1 col-sm-8 col-sm-offset-2 col-md-4 col-md-offset-4">
			<div class="login-panel panel panel-default">
				<div class="panel-heading">PPBMS - Log in</div>
				<div class="panel-body">
					<form role="form" onsubmit="return false;">
						<fieldset>
							<span id="statuslog"></span>
							<div class="form-group">
								<input class="form-control" placeholder="Username" id="elog" name="email" type="email" autofocus="">
							</div>
							<div class="form-group">
								<input class="form-control" placeholder="Password" id="plog" name="password" type="password" value="">
							</div>
							<button id="logBtn" class="btn btn-primary" onclick="login()">Log In</button>
						</fieldset>
					</form>
				</div>
			</div>
		</div><!-- /.col-->
	</div><!-- /.row -->	
	
		

	<script src="js/jquery-1.11.1.min.js"></script>
	<script src="js/bootstrap.min.js"></script>
	<script src="js/chart.min.js"></script>
	<script src="js/chart-data.js"></script>
	<script src="js/easypiechart.js"></script>
	<script src="js/easypiechart-data.js"></script>
	<script src="js/bootstrap-datepicker.js"></script>
	<script src="js/main.js"></script>
    <script src="js/ajax.js"></script>
	<script>
		!function ($) {
			$(document).on("click","ul.nav li.parent > a > span.icon", function(){		  
				$(this).find('em:first').toggleClass("glyphicon-minus");	  
			}); 
			$(".sidebar span.icon").find('em:first').addClass("glyphicon-plus");
		}(window.jQuery);

		$(window).on('resize', function () {
		  if ($(window).width() > 768) $('#sidebar-collapse').collapse('show')
		})
		$(window).on('resize', function () {
		  if ($(window).width() <= 767) $('#sidebar-collapse').collapse('hide')
		})
		function login() {
          var e = _("elog").value;
          var p = _("plog").value;
          var status = _("statuslog");
          status.innerHTML = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i></center>';
          if (e == "" || p == "") {
            status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
          } else {
            _("logBtn").disabled = true;
            status.innerHTML = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i></center>';
            var ajax = ajaxObj("POST", "parsers/login.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
                if (ajax.responseText == "login_failed"){
                  status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Wrong Username or Password.</div>';
                  _("logBtn").disabled = false;
                } else {
				  window.location = "account.php?id="+ajax.responseText+"&dashboard=focus";
                }
              }
            }
            ajax.send("e="+e+"&p="+p);
          }
        }
	</script>	
</body>

</html>
