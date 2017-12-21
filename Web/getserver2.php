<?php
	
	$servername = "35.198.52.148";
	$username = "elusuariodeinmortaljic";
	$password = "1q2w3e4rzxCtgbuhby123456";
	$dbname = "inmortalao";

	$conn = new mysqli($servername, $username, $password, $dbname);
	
	$result = $conn->query("select * from extras where nombre='online'");
	
	if ($result->num_rows > 0)
	{
		$row = $result->fetch_assoc();
	}
	
	$online = $row["valor"];
	
	echo "35.198.52.148;7777;InmortalAO ( ".$online."/1000 );0";
?>