<?php
$db_conn = mysqli_connect("localhost", "root", "", "ppbms");
// Evaluate the connection
if (mysqli_connect_errno()) {
	echo mysqli_connect_errno();
	exit();
}
?>