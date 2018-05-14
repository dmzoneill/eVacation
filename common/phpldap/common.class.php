<?php

class Common
{
	public $current_user = null;
	protected $container = array();
	protected $docroot = null;
	protected $userimages = array();

	public function __construct( $ldap )
	{
		$parts = explode( "/" , str_replace( "\\" , "/" , __FILE__ ) );
		array_pop( $parts );
		array_pop( $parts );
		$path = implode( "/" , $parts );		
		$this->docroot = $path . "/";
				
		$this->prepare( $ldap );
		$this->userimages = glob( $this->docroot . "images/users/*.jpg" );
	}

	public function __get( $key ) 
	{		
		if( !array_key_exists( $key , $this->container ) )
		{
			return false;
		}

		return $this->container[ $key ];
	}
	
	private function prepare( $ldap )
	{
		if( isset( $_SERVER[ 'PHP_AUTH_USER' ] ) || isset( $_SERVER[ 'AUTH_USER' ] ) )
		{
			$authuser = isset( $_SERVER[ 'PHP_AUTH_USER' ] ) ? $_SERVER[ 'PHP_AUTH_USER' ] : $_SERVER[ 'AUTH_USER' ];
			
			if( stristr( $authuser , "\\" ) )
			{
				$authuser = substr( $authuser , strpos( $authuser , "\\" ) + 1 );
			}
		}
				
		$this->current_user = $ldap->getldapuser( $authuser );
	}

	public function getuserimage( $idsid , $ldap )
	{		
		if( in_array( $this->docroot . "images/users/" . $idsid . ".jpg" , $this->userimages ) )
		{
			return "images/users/" . $idsid . ".jpg";
		}
		else
		{
			return "images/person.png";
		}
	}
}
