<?php


	$servername = "35.198.52.148";
	//$servername = "localhost";
	$username = "root";
	$password = "123456";
	$dbname = "inmortalao";

	
	// Create connection
	$conn = new mysqli($servername, $username, $password, $dbname);

	// Check connection
	if ($conn->connect_error) {
		die("Connection failed: " . $conn->connect_error);
	} 

	if ($_POST) {
		
		foreach($_POST as $key => $value) {
			$sql = "Update dats set valor='" . $value. "' where nombre='". $key."'";
			$result = $conn->query($sql);
		}
	}
	
	
	$sql = "SELECT valor FROM dats where nombre='Hechizos' limit 1";
	$result = $conn->query($sql);
	
	echo "<form action='' method='post'>";
	
	if ($result->num_rows > 0) 
	{
		$row = $result->fetch_assoc();
		
		echo "<textarea name='Hechizos' cols='100' rows='40'>".$row["valor"]."</textarea>";
		
	}
	
	echo "<input type='submit'  style='width:100%' value='Guardar' />";
	echo "</form>";
	
?>