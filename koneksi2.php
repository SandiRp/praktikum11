<?php
$servername = "localhost";
$username = "root";
$password = "";
$dbname = "web";
$conn=mysqli_connect($servername, $username, $password, $dbname) or die (mysqli_error());
if (!$conn) {
	die("Connection failed ".mysqli_connect_error());
	}