<html>
    <head>
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
        <link rel="icon" href="imgs/InmortalAO.ico">
        <title>InmortalAO - El Regreso</title>
    
        <style type="text/css">
            body { background:top center url('imgs/pared.jpg');}
            .diome { text-align:center; }
            .info { 
                background: #361A0F;
                color: #B29E79;
                padding: 5px;
                margin-top: 10px;
                margin-bottom:20px;
            }
            h2 {
                background: no-repeat center bottom url('imgs/linea-simple.png');
                padding: 0px 0px 30px 0px;
                margin: 0 0 10px 0;
            }
            .fecha {
                float: right;
            }
            a {
                color: #B29E79;
            }
            a.mas { color:#361A0F }
        </style>
		
		<link rel="stylesheet" type="text/css" media="all" href="styles.css" />
		
    </head>
    <body>
	
	<?php
	include_once('include/users.php');
	//include_once('include/functions.php');
	
	
	$servername = "35.198.52.148";
	//$servername = "localhost";
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
	
	if(isset($_GET))
	{
		if (isset($_GET["code"]))
		{
			if (isset($_GET["cuenta"]))
			{
				$cuenta = $_GET["cuenta"];
				$codigo = $_GET["code"];
				
			
				if ($conn->connect_error) {
					die("Connection failed: " . $conn->connect_error);
				}
				
				$result = $conn->query("select * from cuentas where nombre = '$cuenta' and codigo_verificacion='$codigo' and verificada=0");
				
				if ($result->num_rows > 0)
				{
					$row = $result->fetch_assoc();
					$id_cuenta = $row["id"];
					
					$conn->query("update cuentas set verificada=1,codigo_verificacion='' where id=$id_cuenta");
					
					echo '<script language="javascript">alert("Cuenta activada satisfactoriamente.");</script>';	
				}
				
				$conn->close();
				
			}
		}
	}
	
	if(isset($_POST)){
		
		if(isset($_POST['formu']) and 'Crear Cuenta' === $_POST['formu']){
			
			if(isset($_POST['your_username'],$_POST['your_email'],$_POST['your_password'],$_POST['your_password_2'])){
				
				if($_POST['your_username']!="")
				{
					if($_POST['your_email']!="")
					{
						if($_POST['your_password']!="")
						{
							if($_POST['your_password']==$_POST['your_password_2'])
							{
							
								if ($conn->connect_error) {
									die("Connection failed: " . $conn->connect_error);
								} 
								
								$nombreCuenta = $_POST['your_username'];
								$password = $_POST['your_password'];
								$email = $_POST['your_email'];
								
								$result = $conn->query("select * from cuentas where nombre = '$nombreCuenta'");
								
								if ($result->num_rows > 0)
								{
									echo '<script language="javascript">alert("La cuenta ya existe.");</script>';
								}else
								{
									
									$result = $conn->query("select * from cuentas where mail = '$email'");
									
									if ($result->num_rows > 0)
									{
										echo '<script language="javascript">alert("El mail ya esta registrado.");</script>';
									}else
									{
										$code = aleatoria(20);
										$to = $_POST['your_email'];
										$subject = "InmortalAO - Creacion de Cuenta";
										$txt = '<html><body>';
										$txt .= 'Hola '.ucwords($_POST['your_username']).':<br/>';
										$txt .='<br/>Has creado una Cuenta en <strong>InmortalAO</strong>. Para completar la activación haz clic en el siguiente link:';
										$txt .='<br/><br/>';
										$txt .='<a href="http://inmortalao.com.ar/index.php?code='.$code.'&cuenta='.$nombreCuenta.'">Link de activacion de cuenta</a>';
										$txt .='</body></html>';
										
										$headers  = 'MIME-Version: 1.0' . "\r\n";
										$headers .='Content-type: text/html; charset=iso-8859-1' . "\r\n";
										$headers .= "From: no-reply@inmortalao.com.ar" . "\r\n";
										
										$reg = $conn->query("insert into cuentas (nombre,password,mail,verificada,codigo_verificacion) values ('$nombreCuenta','$password','$email',0,'$code')");
										
										mail($to,$subject,$txt,$headers);
										
										$_POST = array();
										
										echo '<script language="javascript">alert("Se ha enviado un mail para la activacion de la cuenta.");</script>'; 
									}
								}
								
								$conn->close();
						
							}else
							{
								echo '<script language="javascript">alert("Las passwords deben coincidir.");</script>'; 
							}
						}else
						{
							echo '<script language="javascript">alert("Debe completar la password de la cuenta.");</script>'; 
						}
						
					}else
					{
						echo '<script language="javascript">alert("Debe completar el email de la cuenta.");</script>'; 
					}
				}else
				{
					echo '<script language="javascript">alert("Debe completar el nombre de la cuenta.");</script>'; 
				}				
			}
		}
    }     
	?>
		 
		 
		<div class="content">
		
			<div style="margin-left: 0%;">
				<a href="https://www.facebook.com/InmortalAO/" target="_blank"><img src="imgs/facebook.ico"></a>
				<a href="https://www.youtube.com/channel/UCKT8ksx1cZqmHSTrrvH_VXQ" target="_blank"><img src="imgs/youtube.jpg"></a>
			</div>		
		
			<div style="color: #b79855;font-size: 20;margin-left: 37%;margin-top: 0%;">
				<label>Usuarios Online: <?php echo $online; ?> / 1000</label>
			</div>
		

			<div class="body" style="margin-left: 16%;margin-top: 5%;">
		
			<div class="box" width="10%" height="10%" style="margin-left: 30%">
				<a href="https://sourceforge.net/projects/inmortalaoreloaded/files/latest/download"><img src="imgs/new/bot-descargar.png"></a>
			</div>
		
		<div class="box" style="margin-left: 10%;">
        <p>Complete el formulario para crear una cuenta gratuita. Los detalles del registro se les enviaran para confirmacion, asegurece de colocar un Email valido. Una vez registrado podra ingresar al juego para comenzar su aventura.</p>
        
        <form action="" method="post" class="loginform" name="registerform" id="registerform">
        
            <p>
                <label>Cuenta:</label>
                <input tabindex="1" type="text" class="text" name="your_username" id="your_username" value="<?php if(isset($_POST['your_username'])){echo $_POST['your_username'];} ?>">
            </p>
            
            <p>
                <label>Email:</label>
                <input tabindex="2" type="text" class="text" name="your_email" id="your_email" value="<?php if(isset($_POST['your_email'])){echo $_POST['your_email'];} ?>">
            </p>
            
            <p>
                <label>Contrase&ntilde;a:</label>
                <input tabindex="3" type="password" class="text" name="your_password" id="your_password" value="<?php if(isset($_POST['your_password'])){echo $_POST['your_password'];} ?>">
            </p>
            
            <p>
                <label>Repita la Contrase&ntilde;a:</label>
                <input tabindex="4" type="password" class="text" name="your_password_2" id="your_password_2" value="<?php if(isset($_POST['your_password_2'])){echo $_POST['your_password_2'];} ?>">
            </p>
            
            <div id="checksave" style="margin-left: 10%;">
            
            <p class="submit">
                <input tabindex="6" type="image" src="imgs/new/procesar.png" name="register" id="wp-submit" value="Crear Cuenta">
                <input type="hidden" name="formu" value="Crear Cuenta">
            </p>
            
            </div>
        
        </form>
        
        <!-- autofocus the field -->
        <script type="text/javascript">try{document.getElementById('your_username').focus();}catch(e){}</script>
    </div>
    </div>    
		</div>
    
    </body>
</html>
		
