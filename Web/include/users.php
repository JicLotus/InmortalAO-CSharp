<?php
 
	
	function login($user, $pass){
		global $c;
		
		$user = $user;
		$pass = $pass;		
		
		$reg = mysqli_query ('SELECT * FROM cuentas WHERE (nombre = LCASE("'.$user.'")) LIMIT 1', $c['db']['game_con']);
		
		
		echo "<script>alert('". mysqli_num_rows($reg) . "')</script>";

		if(!$reg) return -1;
		
		if (mysql_num_rows($reg) == 0) return 0;
		
		if($fila = mysql_fetch_assoc($reg)){

			if($fila['Password'] === $pass){
				
				$_SESSION['user']['name'] = $fila['nombre'];
				$_SESSION['user']['ban'] = $fila['ban'];
				$_SESSION['user']['id'] = $fila['id'];
				$_SESSION['user']['online'] = $fila['Online'];
				$_SESSION['user']['email'] = $fila['mail'];
				$_SESSION['user']['pass'] = $fila['Password'];
				$_SESSION['user']['numpjs'] = $fila['numpjs'];
				$_SESSION['user']['donador'] = $fila['Donador'];
				$_SESSION['user']['creditos'] = $fila['creditos'];
				$_SESSION['user']['verificada'] = $fila['verificada'];
				$_SESSION['user']['identificado'] = true;
				
				// Seguridad
				$_SESSION['user']['REMOTE_ADDR'] = $_SERVER['REMOTE_ADDR'];
				$_SESSION['user']['HTTP_USER_AGENT'] = $_SERVER['HTTP_USER_AGENT'];
				
				return $fila['id'];
			}
			
			cuenta_suceso($fila['id'], $fila['nombre'].': Se intentó ingresar a su cuenta con "'.$pass.'" desde la ip '.getRealIp().'.');
			
			return -1;
		}
	}
	
	function logout(){
		$_SESSION['user'] = array();
      	session_destroy ();
	}
	
	function is_loged(){
		if( (isset($_SESSION['user']['identificado']) and $_SESSION['user']['identificado'] == true) and 
			($_SESSION['user']['REMOTE_ADDR'] == $_SERVER['REMOTE_ADDR']) and 
			($_SESSION['user']['HTTP_USER_AGENT'] == $_SERVER['HTTP_USER_AGENT'])
		) return true;
		else return false;
	}
	
	function registrar_cuenta($usuario, $pass, $email, $code){
		global $c;

		$usuario = mysqli_real_escape_string($usuario);
		$pass = mysqli_real_escape_string($pass);
		$email = mysqli_real_escape_string($email);

		$reg = mysqli_query ('INSERT INTO cuentas 
			(`id`, `nombre`, `Password`, `mail`, `ban`, `Online`, `numpjs`, `Donador`, `creditos`, `betatest`, `verificada`) VALUES 
			(NULL, LCASE("'.$usuario.'"), "'.$pass.'", LCASE("'.$email.'"), "0", "0", "0", "0", "0", "0", "'.$code.'")
		', $c['db']['game_con']);
		
		if(!$reg) return -1;

		return mysqli_insert_id($c['db']['game_con']);
	}
	
	function borrar_cuenta($id){
		global $c;
		
		$pjs = get_personajes_cuenta($id, 20);
		
		foreach($pjs as $pj){
			user_del($pj['IndexPJ']);
		}
		
		$reg = mysql_query('DELETE FROM cuentas WHERE id = '.$id.' LIMIT 1', $c['db']['game_con']);
		
	}
	
	function cuenta_code($user, $code){
		global $c;
		
		$user = mysqli_real_escape_string($user);
		$code = mysqli_real_escape_string($code);		
		
		$reg = mysql_query ('SELECT * FROM cuentas WHERE (nombre = LCASE("'.$user.'")) LIMIT 1', $c['db']['game_con']);

		if(!$reg) return -1;
		
		if (mysql_num_rows($reg) == 0) return 0;
		
		if($fila = mysql_fetch_assoc($reg)){
		
			if($fila['verificada'] === $code or $fila['verificada'] === 'Si'){
					
				$reg = mysql_query ('UPDATE cuentas SET verificada = "Si" WHERE (nombre = LCASE("'.$user.'")) LIMIT 1', $c['db']['game_con']);
					
				$_SESSION['user']['verificada'] = 'Si';
				
				return 1;
			}
		}
		
		return -1;
	}
	
	function cuenta_code_set($id, $email, $code){
		global $c;

		$email = mysql_real_escape_string($email);
		
		$reg = mysql_query ('UPDATE cuentas SET verificada = "'.$code.'", mail = "'.$email.'" WHERE (id = "'.$id.'") LIMIT 1', $c['db']['game_con']);
		
		$_SESSION['user']['email'] = $email;
		$_SESSION['user']['verificada'] = $code;
		
		return mysql_affected_rows();
	}
	
	function cuenta_exist($cuenta, $by = 'nombre'){
		global $c;
		
		$cuenta = mysqli_real_escape_string($cuenta);
		
		$reg = mysql_query ('SELECT * FROM cuentas WHERE ('.$by.' = LCASE("'.$cuenta.'")) LIMIT 1', $c['db']['game_con']);

		if( mysql_num_rows($reg) === 0) return -1;
		else {
			if($fila = mysql_fetch_assoc($reg)){
			
				return $fila['IndexPJ']; 
			} else { return 0; }
		}
	}
	
	/* Function get_personajes_cuenta
	*  Retorna:
	*    -1 si hay algun error
	*    0 si no hay resultados
	*    array de datos si tiene exito
	*/
	function get_personajes_cuenta($id, $limite=10){
		global $c;
		$temp = array();
		
		
		
		$reg = mysql_query ('SELECT * FROM charflags,charstats 
			WHERE (charflags.id = '.$id.')
			AND charflags.IndexPJ = charstats.IndexPJ 
		LIMIT '.$limite, $c['db']['game_con']);

		if(!$reg) return -1;
		else {
			if (mysql_num_rows($reg) == 0) return 0;
			
			while($fila = mysql_fetch_assoc ($reg)){
				$temp[] = $fila;
			}
			mysql_free_result($reg);
			return $temp; 
		}
	}
	
	
	function get_personaje_online($id){
		global $c;
		$temp = array();
		
		$reg = mysql_query ('SELECT * FROM charflags
			WHERE (charflags.IndexPJ = '.$id.') 
		LIMIT 1', $c['db']['game_con']);
		
		if( mysql_num_rows($reg) > 1) return -1;
		else {
			if($fila = mysql_fetch_assoc ($reg)){
				return $fila['Online']; 
			} else { return -2; }
		}
	}
	
	function cuenta_setPass($id, $passNew) {
		global $c;

		$passNew = mysql_real_escape_string($passNew);
		
		$reg = mysql_query ('UPDATE cuentas SET Password = "'.$passNew.'" WHERE id = '.$id.' LIMIT 1', $c['db']['game_con']);
		return mysql_affected_rows();
	}
	
	/* Funcion pj_modName
	* Modifica el nombre de un pj comprobando que sea el propietario de dicho pj
	* y verificando que el pj no este online y no este ban
	*
	* Retorna:
	*	Num de filas afectadas
	*/
	function pj_modName($id, $pj, $nombre) {
		global $c;

		$nombre = mysql_real_escape_string($nombre);

		$reg = mysql_query ('UPDATE charflags SET  Nombre =  "'.$nombre.'" WHERE  IndexPJ ='.$pj.' AND id='.$id.' AND Ban=0 AND Online=0 LIMIT 1 ;', $c['db']['game_con']);
		return mysql_affected_rows();
	}
	
	function cuenta_datos($id){
		global $c;

		$reg = mysql_query ('SELECT * FROM cuentas WHERE id='.$id.' LIMIT 1', $c['db']['game_con']);

		if( mysql_num_rows($reg) > 1) return -1;
		else {
			if($fila = mysql_fetch_assoc ($reg)){
				return $fila; 
			} else { return 0; }
		}
	}
	
	
	/* Function pj_borrar
	*  Retorna:
	*    -1 si hay algun error
	*    0 si no se borro el pj
	*    1 si se borro correctamente
	*/
	function pj_borrar($id, $pj){
		global $c;
		
		$pj = mysql_real_escape_string($pj);
		
		$reg = mysql_query ('SELECT * FROM charflags 
			WHERE (charflags.id = '.$id.')
			AND (charflags.IndexPJ = '.$pj.') 
		LIMIT 1', $c['db']['game_con']);

		if(!$reg) return -1;
		else {
			if (mysql_num_rows($reg) == 0) return 0;
			
			if($fila = mysql_fetch_assoc ($reg)){
				// Borramos el indexPj de todas las tablas
				foreach($c['db']['tablas'] as $tabla) $reg = mysql_query('DELETE FROM '.$tabla.' WHERE IndexPJ = '.$pj.' LIMIT 1', $c['db']['game_con']);
				
				$reg = mysql_query ('UPDATE cuentas SET numpjs = numpjs-1 WHERE id = '.$id.' LIMIT 1', $c['db']['game_con']);
				return mysql_affected_rows();
			}
			mysql_free_result($reg);
			return 1; 
		}
	}
	
	function is_donadorImg($dona, $text = true){
		global $c;
		
		if(1 <= $dona){
			echo '<img src="',$c['site']['url'],'imgs/dona.png" class="star" />';
			if($text) echo ' Si. Te quedan '.$_SESSION['user']['donador'].' d&iacute;as.';
		} else {
			echo '<img src="',$c['site']['url'],'imgs/no-dona.png" class="star" />';
			if($text) echo ' No (<a href="',$c['site']['url'],'panel-de-usuario.php?ac=mercado-ao">Convertite en donador</a>)';
		}
	}
	
	function cuenta_modCreditos($id, $creditos) {
		global $c;

		$reg = mysql_query ('UPDATE cuentas SET creditos =  creditos + "'.$creditos.'" WHERE id = '.$id.' LIMIT 1', $c['db']['game_con']);
		//echo 'UPDATE cuentas SET creditos =  creditos + "'.$creditos.'" WHERE id = '.$id.' LIMIT 1';
		
		return mysql_affected_rows();
	}
	
	function cuenta_modDonador($id, $tiempo) {
		global $c;

		$reg = mysql_query ('UPDATE cuentas SET Donador =  Donador + "'.$tiempo.'" WHERE id = '.$id.' LIMIT 1', $c['db']['game_con']);
		//echo 'UPDATE cuentas SET Donador =  Donador + "'.$tiempo.'" WHERE id = '.$id.' LIMIT 1';
		return mysql_affected_rows();
	}
	
	function pj_modonline($id, $indexPJ, $estado) {
		global $c;

		$reg = mysql_query ('UPDATE charflags SET Online = "'.$estado.'" WHERE id = '.$id.' AND IndexPJ = '.$indexPJ.' LIMIT 1', $c['db']['game_con']);
		//echo 'UPDATE cuentas SET Donador =  Donador + "'.$tiempo.'" WHERE id = '.$id.' LIMIT 1';
		return mysql_affected_rows();
	}
	
	function pj_modskill($indexPJ, $skill, $cant) {
		global $c;

		$reg = mysql_query ('UPDATE charskills SET SK'.$skill.' = "'.$cant.'" WHERE IndexPJ = '.$indexPJ.' LIMIT 1', $c['db']['game_con']);
		//echo 'UPDATE charskills SET SK'.$skill.' = "'.$cant.'" WHERE IndexPJ = '.$indexPJ.' LIMIT 1';
		return mysql_affected_rows();
	}
	
	function pj_correoAdd($indexPJ, $msg, $de, $cant, $item){
		global $c;
		
		$msg = mysql_real_escape_string($msg);
		
		$reg = mysql_query ('INSERT INTO charcorreo 
			(`Idmsj`, `IndexPJ`, `Mensaje`, `De`, `Cantidad`, `Item`) VALUES 
			(NULL, "'.$indexPJ.'", "'.$msg.'", "'.$de.'", "'.$cant.'", "'.$item.'")
		', $c['db']['game_con']);
		
		logs('INSERT INTO charcorreo (`Idmsj`, `IndexPJ`, `Mensaje`, `De`, `Cantidad`, `Item`) VALUES (NULL, "'.$indexPJ.'", "'.$msg.'", "'.$de.'", "'.$cant.'", "'.$item.'")');
		
		if(!$reg) return -1;

		return mysql_insert_id($c['db']['game_con']);
	}
	
	function logs($text = ''){
		$fp = fopen('logos.txt',"a");
			fwrite($fp, $text. PHP_EOL);
		fclose($fp);
	}
	
	
	function faccion($dato){
		if($dato['Ciudadano']) return 1;
		elseif($dato['Republicano']) return 3;
		elseif($dato['Renegado']) return 2;
		elseif($dato['EjercitoReal']) return 6;
		elseif($dato['EjercitoCaos']) return 5;
		elseif($dato['EjercitoMili']) return 7;	
		
		// error
		return -1;	
	}
	
	
	
	
	
	
	
	
	
	function user_pj($indexPJ){
		global $c;

		$reg = mysql_query ('SELECT * FROM charflags WHERE (indexPJ = '.$indexPJ.')', $c['db']['game_con']);
		return mysql_affected_rows();
	}
	
	function user_get($indexPJ, $tabla){
		global $c;

		$reg = mysql_query ('SELECT * FROM '.$tabla.' WHERE (indexPJ = '.$indexPJ.') LIMIT 1', $c['db']['game_con']);

		if( mysql_num_rows($reg) > 1) return -1;
		else {
			if($fila = mysql_fetch_assoc ($reg)){
				return $fila; 
			} else { return 0; }
		}
	}


	function user_del($indexPJ){
		global $c;

		foreach($c['db']['tablas'] as $tabla) $reg = mysql_query('DELETE FROM '.$tabla.' WHERE IndexPJ = '.$indexPJ.' LIMIT 1', $c['db']['game_con']);
	}
	
	/* Function pj_exist
	*  Retorna:
	*    -1 si hay algun error
	*    0 si el usuario es incorrecto
	*    IdexPJ si tiene exito
	*/
	function pj_exist($pj){
		global $c;
		
		$pj = mysql_real_escape_string($pj);
		
		$reg = mysql_query ('SELECT * FROM charflags WHERE (Nombre = LCASE("'.$pj.'")) LIMIT 1', $c['db']['game_con']);

		if( mysql_num_rows($reg) > 1) return -1;
		else {
			if($fila = mysql_fetch_assoc($reg)){
			
				return $fila['IndexPJ']; 
			} else { return 0; }
		}
	}
	
	/* Function user_compruevaPjEmail
	*  Retorna:
	*    -1 si hay algun error
	*    0 si el e-mail y el usuario son incorrectos
	*    IdexPJ si tiene exito
	*/
	function user_compruevaEmail($id, $email) {
		global $c;
		
		$email = mysql_real_escape_string($email);
		
		$reg = mysql_query ('SELECT *,LCASE(Email) as mail FROM cuentas WHERE (id = '.$id.' AND mail = LCASE("'.$email.'"))', $c['db']['game_con']);
		
		if( mysql_num_rows($reg) > 1) return -1;
		else {
			if($fila = mysql_fetch_assoc ($reg)){
				return $fila['mail']; 
			} else { return 0; }
		}
	}
	
	
	function user_publicPass($id, $pj){
		global $c;
		
		return(md5($pj.$c['site']['random']).'-'.$id.'-'.md5(rand(0,1000)));
	}
	
	function aleatoria($caracteres = 10){
		//cambiar password a una nueva aleatoria
		$abc = array('a','b','c','d','e','f','g','h','i','j','k','l','m','n','o','p','q','r','s','t','u','v','w','x','y','z');
		$max = count($abc)-1;
		
		$passNew = '';
		for($i=0; $i < $caracteres; $i++){
			if(rand(0,2)) $passNew .= $abc[rand(0,$max)];
			else $passNew .= rand(0,9);
		}
		
		return $passNew;
	}
	
	function server_getExtra($nombre){
		global $c;
		
		$reg = mysql_query ('SELECT * FROM extras WHERE (nombre = "'.$nombre.'") LIMIT 1', $c['db']['game_con']);

		if( mysql_num_rows($reg) > 1) return -1;
		else {
			if($fila = mysql_fetch_assoc($reg)){	
				return $fila['valor']; 
			} else { return 0; }
		}
	}
	
	function server_getCount($tabla){
		global $c;
		
		$reg = mysql_query ('SELECT COUNT(*) as valor FROM '.$tabla, $c['db']['game_con']);

		if( mysql_num_rows($reg) > 1) return -1;
		else {
			if($fila = mysql_fetch_assoc($reg)){	
				return $fila['valor']; 
			} else { return 0; }
		}
	}
?>