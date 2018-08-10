<?php
include_once("includes/loginstatus.php");
// If user is already logged in
if (isset($_SESSION["userid"])) {
  header("location: admin.php?id=".$_SESSION["userid"]."&info=focus");
  exit();
}
?>
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Alab - Sign Up</title>

<link href="css/bootstrap.min.css" rel="stylesheet">
<link href="css/datepicker3.css" rel="stylesheet">
<link href="css/styles.css" rel="stylesheet">
<link href="css/custom/login.css" rel="stylesheet">
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
				<div class="panel-heading">Alab - Sign Up</div>
				<div class="panel-body">
					<form role="form" onsubmit="return false;">
						<fieldset>
							<span id="statusreg"></span>
							<div class="form-group">
								<select class="form-control" id="reg-town">
									<option value="">Town</option>
									<?php 
										$sql = "SELECT * FROM alab_town WHERE town_status = '1'";
										$query = mysqli_query($db_conn, $sql);
										while ($row = mysqli_fetch_array($query, MYSQLI_ASSOC)) {
										  $id = $row["id"];
										  $name = $row["town_name"];
										  echo'<option value="'.$id.'">'.$name.'</option>';
										}

										mysqli_close($db_conn);

									?>
								</select>
							</div>
							<div class="form-group">
								<input class="form-control" placeholder="E-mail" id="reg-email" name="email" type="email">
							</div>
							<div class="form-group">
								<input class="form-control" placeholder="Password" id="reg-pass" name="password" type="password" value="">
							</div>
							<a id="regBtn" class="btn btn-primary" onclick="register()">Create</a><a href="login.php" class="btn btn-default pull-right">Login</a>
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
		function register() {
          var t = _("reg-town").value;
          var e = _("reg-email").value;
          var p = _("reg-pass").value;
          var status = _("statusreg");
          status.innerHTML = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999; margin-bottom: 10px;"></i></center>';
          if (t == "" || e == "" || p == "") {
            status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Incomplete Parameters.</div>';
          } else {
            _("regBtn").disabled = true;
            status.innerHTML = '<center><i class="fa fa-circle-o-notch fa-spin" style="color: #999;"></i></center>';
            var ajax = ajaxObj("POST", "parsers/register.php");
            ajax.onreadystatechange = function() {
              if(ajaxReturn(ajax) == true) {
                if (ajax.responseText == "successinsert"){
                 	window.location = "index.php";
                } else {
					status.innerHTML = '<div style="color: red; margin-bottom: 10px;">Email or Town is already registered.</div>';
                }
              }
            }
            ajax.send("t="+t+"&e="+e+"&p="+p);
          }
        }
	</script>	
</body>

</html>
