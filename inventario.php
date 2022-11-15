<?php
session_start();
require_once $_SERVER['DOCUMENT_ROOT']."/InventarioAuditores/Controllers/inventarios.php";
if(isset($_GET['closeSession'])){
	session_destroy();
	header('Location: http://servermatrixxxb.ddns.net:8181/InventarioAuditores/views/login/index.php');
}else{
	$invController = new InventariosController();
	if($invController->checkSession())
		header('Location: http://servermatrixxxb.ddns.net:8181/InventarioAuditores/views/inventario.php');
	else
		header('Location: http://servermatrixxxb.ddns.net:8181/InventarioAuditores/views/login/index.php');
}