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
			$sql = "Update balance set valor=" . $value. " where nombre='". $key."'";
			$result = $conn->query($sql);
		}
	}
	
	
	$sql = "SELECT * FROM balance";
	$result = $conn->query($sql);

	echo "Comando para regarcar en el servidor: /reloadsini";
	
	echo "<form action='' method='post'>";
	
	if ($result->num_rows > 0) {
		
		
		echo "<table >";
		
		$count = 0.0;
		$cuentaRegistros = 0;
	
		while($row = $result->fetch_assoc()) {
			
			 if ($cuentaRegistros==0) 
			 {
				 echo "<tr></tr><tr></tr><tr></tr>";
				 echo "<tr>";
				 echo "<th>". "Intervalos:" . "</th>";
				 echo "</tr>";
				 echo "<tr></tr><tr></tr><tr></tr>";
			 }
			
			 if ($cuentaRegistros==26) 
			 {
				 echo "<tr></tr><tr></tr><tr></tr>";
				 echo "<tr>";
				 echo "<th>". "Stats:" . "</th>";
				 echo "</tr>";
				 echo "<tr></tr><tr></tr><tr></tr>";
			 }
			
			// if ($cuentaRegistros==56) 
			// {
				// echo "<tr></tr><tr></tr><tr></tr>";
				// echo "<tr>";
				// echo "<th>". "Lucha:" . "</th>";
				// echo "</tr>";
				// echo "<tr></tr><tr></tr><tr></tr>";
			// }
			
			if ($cuentaRegistros==182) 
			{
				echo "<tr></tr><tr></tr><tr></tr>";
				echo "<tr>";
				echo "<th>". "Vida:" . "</th>";
				echo "</tr>";
				echo "<tr></tr><tr></tr><tr></tr>";
			}
			
			if ($cuentaRegistros==200) 
			{
				echo "<tr></tr><tr></tr><tr></tr>";
				echo "<tr>";
				echo "<th>". "Generales:" . "</th>";
				echo "</tr>";
				echo "<tr></tr><tr></tr><tr></tr>";
			}
			
			
			if ($cuentaRegistros==208) 
			{
				echo "<tr></tr><tr></tr><tr></tr>";
				echo "<tr>";
				echo "<th>". "Mana:" . "</th>";
				echo "</tr>";
				echo "<tr></tr><tr></tr><tr></tr>";
			}
			
			
			if ((float)$count / 7.0 == 0.0) {echo "<tr>";}
			
			if ($cuentaRegistros>=182 or ( $cuentaRegistros<56 )  ){
				echo "<th>" . $row["nombre"]. " : " . "<input style='width: 40px;' name='".$row["nombre"] ."' value='".$row["valor"]."' />" . "</th>";
			}
			
			$count++;
			$cuentaRegistros++;
			
			if ((float)$count / 7.0 == 1.0) {
				echo "</tr>";
				$count=0;
			}
			
		}
		
		echo "</table>";
		
		
	} else {
		echo "0 results";
	}
	
	
	echo "<br><br><input type='submit' style='width:100%' value='Guardar' />";
	echo "</form>";
	
	$conn->close();

	
?>








