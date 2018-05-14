<?php

include( "ldapuser.class.php" );
include( "ldap.class.php" );
include( "common.class.php" );

$ldap = new Ldap();

$common = new Common( $ldap );

$user = isset( $_GET[ 'user' ] ) ? $_GET[ 'user' ] : $common->current_user->samaccountname;

$name = "../" . $common->getuserimage( $user , $ldap );
$fp = fopen($name, 'rb');

header( "Content-Type: image/jpeg" );
header( "Content-Length: " . filesize( $name ) );

// dump the picture and stop the script
fpassthru($fp);
exit;