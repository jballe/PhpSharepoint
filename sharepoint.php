<?php

 /**
	\mainpage 	
	
	 @version v0.90 september 2005
	 @author Jesper Balle jb@kfumspejderne.dk

	Forbindelse til sharepoint vha. web services.	  
 */
 
 if (!defined('_SHAREPOINT_LAYER')) {
 	define('_SHAREPOINT_LAYER',1);

	//==============================================================================================	
	// CONSTANT DEFINITIONS
	//==============================================================================================	

	// Pt. anvendes NuSoap http://sourceforge.net/projects/nusoap/
	// Man kan evt. overveje at kigge p� PEARs soap-klienter i stedet
	include_once("/home/wwwsites/generel/shared/nusoap/nusoap.php");

	//==============================================================================================	
	// CLASS SharepointSite
	//==============================================================================================	

  class SharepointSite {
	//
	// PUBLIC VARS
	//
	var $debug = 0;		// Debug level 
					/*	0: intet output
						1: fejl vises
						3: alle funktionskald vises
						4: ved fejl vises ogs� kommunikation 
						5: alle funktionskald vises og ved fejl ogs� kommunikation
'						6: alt kommunikation vises
					*/
	
	//
	// PRIVATE VARS
	//
	var $soapUser = null;			// Hvilken bruger skal forbinde til web service
	var $soapPass = null;			// Password tilh�rende brugeren
	var $sitePath = "/gruppe/gruppenavn";	// relativ sti til sharepoint site
	var $siteHost = "sps.spejdernet.dk	";	// sharepoint server hostname
	var $siteProt = "http";			// protokol til sharepoint site
	var $siteDomain = "SPEJDERNET";		// domain for the site
	
	/**
	 * Konstrukt�r
	 */
	function SharepointSite($site, $user=null, $pass=null) {
		$this->setSite($site);
		$this->soapUser = $user;
		$this->soapPass = $pass;
	}
	
	/**
	 * Angiv adresse for Sharepoint sitet
	 * @param	string $site Url til sitet. Enten som en komplet url (http://sps.spejdernet.dk/gruppe/gruppenavn),
	 			uden protokol (sps.spejdernet.dk/gruppe/gruppenavn) eller relativt (/gruppe/gruppenavn).
	 * @access   public
	 */
	function setSite($site) {
		$site = trim($site);
		if ($pos = strpos($site, "://")) {
			$this->siteProt = trim(substr($site, 0, $pos));
			$site = substr($site, $pos+3);
			$pos = strpos($site, "/");
			$this->setHost(substr($site, 0, $pos));
			$this->sitePath = trim(substr($site, $pos));
		} elseif (substr($site,0,1)!="/") {
			$pos = strpos($site,"/");
			$this->setHost(substr($site, 0, $pos));
			$this->sitePath = trim(substr($site, $pos+1));
		} else {
			$this->sitePath = trim($site);
		}

		if (substr($this->sitePath, -1)=="/")
			$this->sitePath = trim(substr($this->sitePath, 0, strlen($this->sitePath)-1));
	}
	
	/**
	 * S�t servernavn for Sharepoint sitet.
	 * @param	string $host servernavn eks. sps.spejdernet.dk
	 * @access	private
	 */
	function setHost($host) {
		$this->siteHost = trim($host);
		$array = explode(".", $host);
		$this->siteDomain = strtoupper($array[sizeof($array)-2]);
	}
	/**
	 * Den fulde adresse p� sitet (uden afsluttende '/').
	 * @return	string url
	 * @access	private
	 */
	function getUrl() {
		return $this->siteProt."://".trim($this->siteHost).trim($this->sitePath);
	}
	
	//==============================================================================================	
	// LISTEINDHOLD
	//==============================================================================================	
	/**
	 * Returner indholdet af en liste
	 * @access	public
	 */
	function listContents($var1, $var2=null) {
		$this->_debugStart("listContent({$list})");
		$inputfail = "Enten gives et listQuery-objekt, eller en tekststreng med listens id som parameter (evt. suppleret med et heltal for antallet af liste-elementer der m� returneres).";

		// valider og parse input
		switch(gettype($var1)) {
			case "string":	// det er en tekststreng
				$query = new listQuery($var1);
				$query->limitRows($var2);
				break;
			case "object":	// det er m�ske en listQuery
				if(get_class($var1)!="listquery")
					return $this->_debugError(10, $inputfail);
				else $query = $var1;
				break;
			default:
				return $this->_debugError(10, $inputfail);
		}

		// generer foresp�rgsel
		$soapPath = $this->getUrl() . "/_vti_bin/DspSts.asmx";
		$soapAction	= "http://schemas.microsoft.com/sharepoint/dsp/queryRequest";
		$soapHeader	= '<request document="content" method="query" xmlns="http://schemas.microsoft.com/sharepoint/dsp" />'
		            . '  <versions xmlns="http://schemas.microsoft.com/sharepoint/dsp">'
		            . '    <version>1.0</version>'
		             . '  </versions>';
		$soapBody    = $query->_makeBody();

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody, $soapHeader);
		if($result !== null) return $result;

		// UDTR�K OG BEARBEJD RESULTAT
		$dsQueryResponse = $this->result[1]["children"][0]["children"][0];
		if($dsQueryResponse["attrs"]["status"]!="success")
			return $this->_debugError(5, $dsQueryResponse["attrs"]["status"]);

		// Lige lidt s� vi kan finde rundt i resultatet		
		$rows  = $dsQueryResponse["children"][1]["children"];
		$skema = $dsQueryResponse["children"][0];
		$field = $skema["children"][0]["children"][0]["children"][0]["children"][0]["children"][0]["children"];
		// Felttyper
		$fields = array();
              if(is_array($field))
		foreach($field as $fieldtype) {
			$key = $this->_normalizeKey($fieldtype["attrs"]["name"]);
			$row = array();
			$row["displayname"] = $this->_normalizeString($fieldtype["attrs"]["d:displayName"]);
			$row["fieldtype"] = isset($fieldtype["attrs"]["type"]) ? $fieldtype["attrs"]["type"] : null;
			$fields[$key] = $row;
		}
		// Data
		$result = array(); 
              if(is_array($rows))
		foreach($rows as $row) {
			$resultRow = array();
			foreach($row["attrs"] as $key => $oldvalue) {
				$newKey = $this->_normalizeKey($key);
				switch ($fields[$newKey]["fieldtype"]) {
					case "x:int":		$value = $this->_normalizeInt($oldvalue); break;
					case "x:float":	$value = $this->_normalizeDouble($oldvalue); break;
					case "x:boolean":	$value = $this->_normalizeBoolean($oldvalue, true); break;
					case "x:dateTime":	$value = $this->_normalizeDate($oldvalue); break;
					case "x:string":	$value = $this->_normalizeString($oldvalue); break;
					default:		$value = $this->_normalize($oldvalue); break;
				}
				$resultRow[$newKey] = $value;
				$resultRow[ $fields[$newKey]["displayname"] ] = $value;
			}
			$result[] = $resultRow;
		}

		// Returner
		$this->_debugEnd();
		return $result;
	}
	
	/**
	 * Inds�t eller rediger listeelementer
	 * @param	string $list GUID for den p�g�ldende liste.
	 * @param	array $values Oplysningerne som skal inds�ttes. Feltnavnet som key og v�rdien som value.
	 * @param	int $id hvis man redigerer et element, skrives dets ID her. Ellers null (eller undlad parametren).
	 * @access	public
	 */
	function listItems($list, $values, $id=null) {
		$this->_debugStart("listItems(".substr($list,0,5)."..., {$values}, {$id})");
		
		// valider input
		if(!is_array($values)) 
			return $this->_debugError(10, "\$values skal v�re et array");
		if(is_array($values[0]) && (!is_array($id) && $id!=null)) 
			return $this->_debugError(10, "ved redigering af flere elementer skal disses id'er angives i et array");
		if(is_array($id)) if(sizeof($values)!=sizeof($id)) 
			return $this->_debugError(10, "ved redigering af flere elementer skal alle deres tilh�rende id'er angives i et array");

		// Opbyg foresp�rgsel
		$soapPath   = $this->getUrl()."/_vti_bin/Lists.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/UpdateListItems";
		
		$br = "\n";
		$soapBody   = '<UpdateListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/">'.$br
		            . '  <listName>'.$list.'</listName>'.$br
		            . '  <updates>'.$br
		            . '    <Batch OnError="Continue" >'.$br;
		if(is_array($values[0])) {
			// Flere elementer samtidigt
			for($i=0; $i<sizeof($values); $i++) {
			  if(is_array($thisvalue = $values[$i])) {
				if(is_array($id))
					$soapBody .= '      <Method ID="'.($i+1).'" Cmd="Update">'.$br
					           . '        <Field Name="ID">'.$id[$i].'</Field>'.$br;
				else
					$soapBody .= '      <Method ID="'.($i+1).'" Cmd="New">'.$br;
				foreach($thisvalue as $key => $value) {
					$soapBody .= '        <Field Name="'.$key.'">';
					if($value === true)		$soapBody .= "1";
					else if($value === false)	$soapBody .= "0";
					else 				$soapBody .= stripslashes($value);
					$soapBody .= '</Field>'.$br;
				}
				$soapBody	.= '      </Method>'.$br;
			  }
			}
		} else {
			// Kun et enkelt element
			if($id != null)
				$soapBody .= '      <Method ID="1" Cmd="Update">'.$br
				          . '        <Field Name="ID">'.$id.'</Field>'.$br;
			else
				$soapBody .= '      <Method ID="1" Cmd="New">'.$br;
			
			foreach($values as $key => $value) {
				$soapBody .= '        <Field Name="'.$key.'">';
				if($value === true)		$soapBody .= "1";
				else if($value === false)	$soapBody .= "0";
				else 				$soapBody .= stripslashes($value);
				$soapBody .= '</Field>'.$br;
			}
			$soapBody	.= '      </Method>'.$br;
		}
		$soapBody .= '    </Batch>'.$br
		           . '  </updates>'.$br
		           . '</UpdateListItems>';
		
		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result !== null) return $result;		

		// Udtr�k resultatet
		$results = $this->result[0]["children"][0]["children"][0]["children"][0]["children"];
		$result = array();
		foreach($results as $elementResult) {
			$resultRow = array();
			$resultRow["errorcode"] = $elementResult["children"][0]["chardata"];
			$resultRow["result"] = ($resultRow["errorcode"]=="0x00000000") ? true : false;
			$row = array();
			foreach($elementResult["children"][ sizeof($elementResult["children"])-1 ]["attrs"] as $key => $value)
				if($key != "xmlns:z")
				$row[ $this->_normalizeKey($key) ] = $this->_normalize($value);
			$resultRow["row"] = $row;
			$result[] = $resultRow;
		}
		if(sizeof($result)==1)
			$result = $result[0];

		// returner
		$this->_debugEnd();
		return $result;
	}
	
	//==============================================================================================	
	// VEDH�FTEDE FILER
	//==============================================================================================	
	/**
	 * Returner vedh�ftede filer fra et listeelement
	 * @param	string $list GUID for den p�g�ldende liste
	 * @param	int $id Det p�g�ldende elements ID.
	 * @access	public
	 */
	function listItemAttachments($list, $id) {
		$this->_debugStart("getAttachments(".substr($list,0,7)."..., {$id})");

		// Opbyg foresp�rgsel
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/GetAttachmentCollection";
		$soapBody = "<GetAttachmentCollection xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\">\n"
		          . "  <listName>{$list}</listName>\n"
		          . "  <listItemID>{$id}</listItemID>\n"
		          . "</GetAttachmentCollection>"; 
		$soapPath = $this->getUrl()."/_vti_bin/Lists.asmx";

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result != null) return $result;

		// Udtr�k resultatet
		$result = array();
		$nodes = $this->result[0]["children"][0]["children"][0]["children"][0]["children"];
		if(is_array($nodes))
		foreach($nodes as $row)
			$result[] = $row["chardata"];
	
		// Returner
		$this->_debugEnd();
		return $result;
	}

	//==============================================================================================	
	// LISTER
	//==============================================================================================	
	
	/**
	 * Oplysninger om samtlige lister og dokumentbiblioteker p� sitet
	 * @access	public
	 */
	function listCollection() {
		$this->_debugStart("getListCollection()");

		// Opbyg foresp�rgsel
		$soapPath = $this->getUrl()."/_vti_bin/Lists.asmx";
		$soapAction	= "http://schemas.microsoft.com/sharepoint/soap/GetListCollection";
		$soapBody	= "<GetListCollection xmlns=\"http://schemas.microsoft.com/sharepoint/soap/\" />";
		
		// Udf�r foresp�rgsel
		$this->_doCall($soapPath, $soapAction, $soapBody);
		if(!is_array($this->result)) return $this->result;		

		// Udtr�k resultatet
		$result = array();
		$lists = $this->result[0]["children"][0]["children"][0]["children"][0]["children"];
		foreach($lists as $row) {
			$list = array();
			foreach($row["attrs"] as $key => $value)
				$list[ $this->_normalizeKey($key)] = $this->_normalize($value);
			$result[] = $list;
		}
		$this->_debugEnd();
		return $result;
	}

	/**
	 * Alle oplysninger om en given liste
	 * @param	string $listName Nanvet p� listen
	 * @access	public
	 */
	function getList($listName) {
		$this->_debugStart("getLists({$listName})");

		// Opbyg foresp�rgsel
		$soapPath = $this->getUrl()."/_vti_bin/Lists.asmx";
		$soapAction	= "http://schemas.microsoft.com/sharepoint/soap/GetList";	
		$soapBody	= '<GetList xmlns="http://schemas.microsoft.com/sharepoint/soap/">'.$br
		            . '  <listName>'.$listName.'</listName>'.$br
		            . '</GetList>';

		// Udf�r foresp�rgsel
		var_dump($this->_doCall($soapPath, $soapAction, $soapBody, null));
		// Returner
		$this->_debugEnd();
		return $result;
	}

	function listViewCollection($listGuid) {
		$this->_debugStart("viewCollection('{$listGuid}')");

		// Opbyg foresp�rgsel
		$soapAction	= "http://schemas.microsoft.com/sharepoint/soap/GetViewCollection";
		$soapBody	= '<GetViewCollection xmlns="http://schemas.microsoft.com/sharepoint/soap/">'.$br
				. '  <listName>'.$listGuid.'</listName>'.$br
				. '</GetViewCollection>';
		$soapPath = $this->getUrl()."/_vti_bin/Views.asmx";

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result != null) return $result;

		// Udtr�k resultatet
		$result = array();
		$nodes = $this->result[0]["children"][0]["children"][0]["children"][0]["children"];
		foreach($nodes as $row) {
			$view = array();
			foreach($row["attrs"] as $key=>$value)
				$view[$this->_normalizeKey($key)] = $this->_normalize($value);
			$result[] = $view;
		}

		// Returner
		$this->_debugEnd();
		return $result;
	}

	//==============================================================================================	
	// BILLEDER
	//==============================================================================================	
	function uploadImageFile($libraryName, $filename, $picturePath, $folder="", $overwrite=false) {
		$this->_debugStart("uploadImageFile ('".$libraryName."', '{$filename}', '".substr($picturePath,0,25)."', '{$folder}', '{$overwrite}')");

		if(!file_exists($picturePath))
			return $this->_debugError(10, "Den angivne fil ({$picturePath}) findes ikke.");

		$data = file_get_contents($picturePath);
		if(!$data)
			return $this->_debugError(10, "Der kunne ikke indl�ses nogen data fra den angivne fil");

		return $this->uploadImage($libraryName, $filename, base64_encode($data), $folder, $overwrite);
	}
	function uploadImage($libraryName, $filename, $picture_data, $folder="", $overwrite=false) {
		$this->_debugStart("uploadImage ('".$libraryName."', '{$filename}', '".substr($picture_data,0,25)."', '{$folder}', '{$overwrite}')");
	
		// opbyg foresp�rgsel
		$soapPath = $this->getUrl()."/_vti_bin/Imaging.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/ois/Upload";
		$soapBody = "<Upload xmlns=\"http://schemas.microsoft.com/sharepoint/soap/ois/\">\n"
				. "	<strListName>".$libraryName."</strListName>\n"
      				. "	<strFolder>".$folder."</strFolder>\n"
		      		. "	<bytes>".$picture_data."</bytes>\n"
		      		. "	<fileName>".$filename."</fileName>\n"
      				. "	<fOverWriteIfExist>";
		$soapBody .= ($overwrite) ? "true" : "false"; 
		$soapBody 	.="</fOverWriteIfExist>\n"
    				. "</Upload>";

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result !== null) return $result;
    	          
		// Udtr�k resultatet
		$result = array();
		$upload = $this->result[0]["children"][0]["children"][0]["children"][0];
		foreach($upload["attrs"] as $key=>$value)
			if($key!="xmlns")
				$result[ $this->_normalizeKey($key) ] = $this->_normalize($value);

		// returner
		$this->_debugEnd();
		return $result;
	}
	function createImgNewFolder($list, $parentfolder) {
		$this->_debugStart("createImgNewFolder ('{$list}', '{$parentfolder}')");
	
		// opbyg foresp�rgsel
		$soapPath = $this->getUrl()."/_vti_bin/Imaging.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/ois/CreateNewFolder";
		$soapBody = "<CreateNewFolder xmlns=\"http://schemas.microsoft.com/sharepoint/soap/ois/\">\n"
				. "	<strListName>".$list."</strListName>\n"
				. "	<strParentFolder>".$folder."</strParentFolder>\n"
				. "</CreateNewFolder>";

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result !== null) return $result;
    	          
		// Udtr�k resultatet
		$result = $this->result[0]["children"][0]["children"][0]["children"][0]["attrs"]["title"];

		// returner
		$this->_debugEnd();
		return $result;
	}
	function renameImg($list, $folder, $renames) {
		$this->_debugStart("renameImg('{$list}', '{$folder}', '{$renames}')");

		// valider input 
		if(!is_array($renames)) return $this->_debugError(10, "renames skal v�re et array � la 'gammel navn =>  ny navn'");

		// opbyg foresp�rgsel
		$soapPath = $this->getUrl()."/_vti_bin/Imaging.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/ois/Rename";
		$soapBody = '<Rename xmlns="http://schemas.microsoft.com/sharepoint/soap/ois/">'."\n"
			   . '  <strListName>'.$list.'</strListName>'."\n"
			   . '  <strFolder>'.$folder.'</strFolder>'."\n"
			   . '  <request>'."\n"
			   . '   <files>'."\n";
		foreach($renames as $key=>$value)
		$soapBody .="    <file filename=\"{$key}\" newbasename=\"{$value}\"/>\n";
		$soapBody .='   </files>'."\n"
			   . '  </request>'."\n</Rename>";

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result !== null) return $result;
    	          
		// Udtr�k resultatet
		$result = $this->result[0]["children"][0]["children"][0]["children"][0]["children"];
		$retur = array();
		foreach($result as $i)
			$retur[] = $i["attrs"];
		

		// returner
		$this->_debugEnd();
		return $retur;
	}
	function createImgFolder($list, $parentFolder, $folderName) {
		$this->_debugStart("createImgFolder('{$list}', '{$parentFolder}', '{$folderName}')");

		$tmp = $this->createImgNewFolder($list, $parentfolder);
		$result = $this->renameImg($list, $parentFolder, array($tmp=>$folderName));

		// returner
		$this->_debugEnd();
		return $result;	
	}
	function imageFolderContents($listname, $folder) {
		$this->_debugStart("renameImg('{$list}', '{$folder}', '{$renames}')");

		// opbyg foresp�rgsel
		$soapPath = $this->getUrl()."/_vti_bin/Imaging.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/ois/GetListItems";
		$soapBody = '<GetListItems xmlns="http://schemas.microsoft.com/sharepoint/soap/ois/">'."\n"
			   . ' <strListName>'.$listname.'</strListName>'."\n"
			   . ' <strFolder>'.$folder.'</strFolder>'."\n"
			   . '</GetListItems>';

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result !== null) return $result;
    	          
		// Udtr�k resultatet
		$result = $this->result[0]["children"][0]["children"][0]["children"][0]["children"];
		$retur = array();
		foreach($result as $i) {
			$row = array();
			foreach($i["attrs"] as $key=>$value)
				if($key != "xmlns:z")
				$row[ $this->_normalizeKey($key) ] = $this->_normalize($value);
			$retur[] = $row;
		}

		// returner
		$this->_debugEnd();
		return $retur;
	}

	//==============================================================================================	
	// USER INFO
	//==============================================================================================	
	/**
	 * Oplysninger om en given bruger som findes p� sitet
	 * @param	string $username Brugerens login
	 * @access	public
	 */
	function userInfo($username) {
		$this->_debugStart("userInfo({$username})");
		
		// Opbyg foresp�rgsel
		$soapPath = $this->getUrl()."/_vti_bin/UserGroup.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/directory/GetUserInfo";
		$soapBody = '<GetUserInfo xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">'
      	          . '	<userLoginName>'.$this->_domainAdd($username).'</userLoginName>'
    	          . '</GetUserInfo>';
    	          
		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result !== null) return $result;
    	          
		// Udtr�k resultatet
		$result = array();
		$attributes = $this->result[0]["children"][0]["children"][0]["children"][0]["children"][0]["attrs"];
		foreach($attributes as $key=>$value)
			$result[$this->_normalizeKey($key)] = $value;
		$this->_debugEnd();
		return $result;
	}
	
	/**
	 * En brugers profil-oplysninger
	 * @param	string $username Brugerens login
	 * @access	public
	 */
	function userProfile($username) {
		$this->_debugStart("userProfile({$username})");
				
		// Opbyg foresp�rgsel
		$soapAction = "http://microsoft.com/webservices/SharePointPortalServer/UserProfileService/GetUserProfileByName";
		$soapBody = '<GetUserProfileByName xmlns="http://microsoft.com/webservices/SharePointPortalServer/UserProfileService">'
      			.	'  <AccountName>'.$this->_domainRemove($username).'</AccountName>'
    			. '</GetUserProfileByName>';
		$soapPath = "http://sps.spejdernet.dk"."/_vti_bin/UserProfileService.asmx";

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result != null) return $result;

		// Udtr�k resultatet
		$result = array();
		$nodes = $this->result[0]["children"][0]["children"][0]["children"];
		if(is_array($nodes))
			foreach($nodes as $property)
				$result[$property["children"][0]["chardata"]] = $this->_normalize($property["children"][1]["chardata"]);
		// Returner
		$this->_debugEnd();
		return $result;
	}
	
	//==============================================================================================	
	// TV�RG�ENDE WEBSTEDSGRUPPER
	//==============================================================================================
	
	/**
	 * Opret en tv�rg�ende webstedsgruppe p� Sharepoint sitet
	 * @param	string $groupName Navnet p� den nye tv�rg�ende webstedsgruppe
	 * @param	string $ownerIdent Hvem ejer webstedsgruppen - brugernavn eller navn p� anden webstedgruppe
	 * @param	string $ownerType Hvad ejer webstedsgruppen - "group" eller "user"
	 * @param	string $defUserLogin Startmedlem - brugernavn
	 * @param 	string $description Beskrivelse til webstedsgruppen
	 * @access	public
	 */
	function addGroup($groupName, $ownerIdent, $ownerType, $defUserLogin, $description="") {
		$this->_debugStart("addGroup({$groupName}, {$ownerIdent}, {$ownerType}, {$defUserLogin}, {$description})");
		
		// Opbyg foresp�rgsel
		$soapPath = $this->getUrl() . "/_vti_bin/UserGroup.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/directory/AddGroup";
		$soapBody	= "<AddGroup xmlns=\"http://schemas.microsoft.com/sharepoint/soap/directory/\">\n"
					. " <groupName>".$groupName."</groupName>\n"
					. " <ownerIdentifier>".$ownerIdent."</ownerIdentifier>\n"
					. " <ownerType>".$ownerType."</ownerType>\n"
					. " <defaultUserLoginName>".$defUserLogin."</defaultUserLoginName>\n"
					. " <description>".$description."</description>\n"
					. "</AddGroup>";
					
		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if(!($result === null)) return $result;

		// Returner
		$this->_debugEnd();
		return true;
	}
	/**
	 * Hvilke tv�rg�ende webstedsgrupper en bruger er medlem af.
	 * @param	string $username Brugerens login
	 * @access	public
	 */
	function groupCollectionFromUser($username) {
		$this->_debugStart("groupCollectionFromUser({$userName})");
		
		// Opbyg foresp�rgsel
		$soapPath = $this->getUrl() . "/_vti_bin/UserGroup.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/directory/GetGroupCollectionFromUser";
		$soapBody = '<GetGroupCollectionFromUser xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">'
		          . '  <userLoginName>'.$this->_domainAdd($username).'</userLoginName>'
		          . '</GetGroupCollectionFromUser>';
	
		// Udf�r foresp�rgsel
		$this->_doCall($soapPath, $soapAction, $soapBody);
		if(!is_array($this->result)) return $this->result;		

		// Udtr�k resultatet
		$groups = $this->result[0]["children"][0]["children"][0]["children"][0]["children"][0]["children"];
		$result = array();
		foreach($groups as $group) {
			$groupinfo = array();
			foreach($group["attrs"] as $key => $value)
				$groupinfo[$this->_normalizeKey($key)] = $this->_normalize($value);
			$result[$groupinfo["name"]] = $groupinfo;
		}
		
		// Returner
		$this->_debugEnd();
		return $result;
	}

	/**
	 * Tilf�j en bruger til en tv�rg�ende webstedsgruppe
	 * @param	string $groupName Den tv�rg�ende webstedgruppes navn
	 * @param	string $userName  Brugerens navn (displayname)
	 * @param	string $userLogin Brugerens login
	 * @param	string $userEmail Brugerens e-mail adresse
	 * @param	string $userNotes Noter for brugeren
	 * @access	public
	 */
	function addUserToGroup($groupName, $userName, $userLogin, $userEmail="", $userNotes="") {
		$this->_debugStart("addUserToGroup({$groupName}, {$userName}, {$userLogin}, {$userEmail}, {$userNotes})");
		
		// Opbyg foresp�rgsel
		$br = "\n";
		$soapPath = $this->getUrl() . "/_vti_bin/UserGroup.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/directory/AddUserToGroup";
		$soapBody = '<AddUserToGroup xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">'.$br
		          . '   <groupName>'.$groupName.'</groupName>'.$br
		          . '   <userName>'.$userName.'</userName>'.$br
		          . '   <userLoginName>'.$this->_domainAdd($userLogin).'</userLoginName>'.$br
		          . '   <userEmail>'.$userEmail.'</userEmail>'.$br
		          . '   <userNotes>'.$userNotes.'</userNotes>'.$br
		          . '</AddUserToGroup>';

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result !== null) return $resultt;
		
		// Udtr�k resultatet
		$this->_debugEnd();
		return true;
	}
	
	/**
	 * Oplysninger om en given tv�rg�ende webstedsgruppe
	 * @param	string $group
	 * @access	public
	 */
	function groupInfo($group) {
		$this->_debugStart("groupInfo({$group})");
		
		// Opbyg foresp�rgsel
		$soapPath = $this->getUrl()."/_vti_bin/UserGroup.asmx";
		$soapAction = "http://schemas.microsoft.com/sharepoint/soap/directory/GetGroupInfo";
		$soapBody =	'<GetGroupInfo xmlns="http://schemas.microsoft.com/sharepoint/soap/directory/">\n'
	    	   	   .	'  <groupName>'.$group.'</groupName>\n'
	    	 	   . '</GetGroupInfo>';

		// Udf�r foresp�rgsel
		$result = $this->_doCall($soapPath, $soapAction, $soapBody);
		if($result !== null) return $result;
		
		// Udtr�k resultatet
		$result = array();
		$attributes = $this->result[0]["children"][0]["children"][0]["children"][0]["children"][0]["attrs"];
		foreach($attributes as $key => $value)
			$result[strtolower($key)] = $this->_normalize($value);

		// Returner
		$this->_debugEnd();
		return $result;
	}
	
	//==============================================================================================	
	// HELPER FUNCTIONS
	//==============================================================================================	
	/**
	 * Fjern et evt. dom�ne fra et brugernavn
	 * @param	string $username brugernavn der evt. indeholder dom�ne
	 * @return	string brugernavnet uden dom�ne
	 * @access	private
	 */
	function _domainRemove($username) {
		// Fjern evt. dom�ne
		if($pos = strpos($username, "\\"))
			$username = substr($username, $pos+1);
		return $username;
	}
	
	/**
	 * Tilf�j dom�ne til et brugernavn
	 * @param	string $username brugernavn der evt. indeholder dom�ne
	 * @return	string brugernavnet med dom�ne
	 * @access	private
	 */
	function _domainAdd($username) {
		$search = strtolower($this->siteDomain)."\\";
		$pos = strpos(strtolower($username), $search);
		if($pos!==false)
			return $username;
		else
			return $this->siteDomain."\\".$username;
	}
	
	// WEB SERVICE FUNCTIONS
	var $client = null;						// Active soapclient

	/**
	 * Udf�r et webservice kald og s�rger for at svaret parses
	 * @param	string $soapUrl adressen til den p�g�ldende webservice
	 * @param	string $soapAction angiver SOAPAction til webservicen
	 * @param	string $soapBody foresp�rgsel der skal sendes til webservicen
	 * @param	string @soapHeader evt. headere der skal sendes med foresp�rgslen til webservicen
	 * @access	private
	 */
	function _doCall($soapUrl, $soapAction, $soapBody, $soapHeader=null) {
		// Kald web service
		$result = $this->_makeRequest($soapUrl, $soapAction, $soapBody, $soapHeader);
		if(!($result === null)) return $result;

		// Parse response
		$result = $this->_parseResponse();
		return $result;
	}
	
	/**
	 * Udf�r et webservice kald
	 * @param	string $soapUrl adressen til den p�g�ldende webservice
	 * @param	string $soapAction angiver SOAPAction til webservicen
	 * @param	string $soapBody foresp�rgsel der skal sendes til webservicen
	 * @param	string @soapHeader evt. headere der skal sendes med foresp�rgslen til webservicen
	 * @access	private
	 */
	function _makeRequest($soapUrl, $soapAction, $soapBody, $soapHeader=null) {
		// Opret web service klient
		$this->client = new soapclient($soapUrl);
		if ($err = $this->client->getError()) return $this->_debugError(1, $err);

		//Credentials
		if($this->soapUser)
			$this->client->setCredentials($this->soapUser, $this->soapPass);
		
		// Send request
		$msg = $this->client->serializeEnvelope($soapBody, $soapHeader);
		$this->client->send($msg,$soapAction);
		if ($err = $this->client->getError()) return $this->_debugError(2, $err);
		return null;
	}
	/**
	 * Udtr�kker xml fra webservice svaret
	 * @return	string xml
	 * @access	private
	 */
	function _getXmlResponse() {
		return substr($this->client->response, strpos($this->client->response, "<?xml") );
	}
	/**
	 * Parser webservice svar (herunder udtr�kker xml)
	 * @access	private
	 */
	function _parseResponse() {
		$doc = $this->_getXmlResponse();
		$result = $this->_parseXml($doc);
		if ($result !== null) return $result;

		$result = $this->_parseError();
		if($result !== null) return $result;
	}

	function _parseError() {
		if($this->result[0]["children"][0]["name"]=="soap:Fault") {
			// Der opstod en fejl
			$fejl = $this->result[0]["children"][0]["children"];
			$detail = $fejl[2]["children"];

			$errorstring = $this->_normalizeString($fejl[1]["chardata"]);
			if(strlen($detail[0]["chardata"])>0)
				$errorstring .= " - "
					.$this->_normalizeString($detail[0]["chardata"]) // errorstring
					." (".
					$this->_normalizeString($detail[1]["chardata"]) // errorcode
					.")";
			return $this->_debugError(5, $errorstring);
		} else
			return null;		
	}	
	function _parseXml($doc) {
		$parser = new xml2array($doc);
		$result = $parser->parse();
		if(!is_array($result))	return debugError(4, $result); // XML Parser error
		else {
			// vi gider ikke have SOAP:ENVELOPE med...
			$this->result = $result[0]["children"];
			return null;
		}
	}
	
	//==============================================================================================	
	// Normalisering af tegn og typer
	// Tilsyneladende opst�r en fejl med danske tegn/special-tegn i forbindelse med kommunikationen
	// - dette fikses her
	//==============================================================================================	
	var $replaceChars = array ( 
				"”" => "\"",
				"æ" => "�",
				"ø" => "�",
				"å" => "�",
				"Æ" => "�",
				"Ø" => "�",
				"Å" => "�",
				"–" => "-",
				"é" => "�",
				"&#39" => "'",
				"â" => "�",
				"" => "-",
				"É" => "�"
			);
	function _normalizeBack($str) {
		return str_replace(	array_values($this->replaceChars),
					array_keys($this->replaceChars),
					$str
				);
	}
	function _normalize($str) {			
		// datoer
		$val = $this->_normalizeDate($str);
		if($val!==null) {
			return $val;
		}

		// heltal
		$val = $this->_normalizeInt($str);
		if(!($val===null))
			return $val;
			
		// decimaltal
		$val = $this->_normalizeDouble($str);
		if(!($val===null))
			return $val;
			
			
		// Boolean
		$val = $this->_normalizeBoolean($str);
		if(!($val===null))
			return $val;

		// tekst
		return $this->_normalizeString($str);
	}
	function _normalizeString($str) {
		foreach($this->replaceChars as $search => $replace)
			$str = str_replace($search, $replace, $str);
		return $str;
	}
	function _normalizeDate($datestr) {
	$eregStr =
	'([0-9]{4})-'.	// centuries & years CCYY-
	'([0-9]{2})-'.	// months MM-
	'([0-9]{2})'.	// days DD
	'T'.		// separator T
	'([0-9]{2}):'.	// hours hh:
	'([0-9]{2}):'.	// minutes mm:
	'([0-9]{2})';//(\.[0-9]+)?'. // seconds ss.ss...
//	'(Z|[+\-][0-9]{2}:?[0-9]{2})?'; // Z to indicate UTC, -/+HH:MM:SS.SS... for local tz's
	if(ereg($eregStr,$datestr,$regs)){
		// not utc
		if($regs[8] != 'Z'){
			$op = substr($regs[8],0,1);
			$h = substr($regs[8],1,2);
			$m = substr($regs[8],strlen($regs[8])-2,2);
			if($op == '-'){
				$regs[4] = $regs[4] + $h;
				$regs[5] = $regs[5] + $m;
			} elseif($op == '+'){
				$regs[4] = $regs[4] - $h;
				$regs[5] = $regs[5] - $m;
			}
		}
//		$result = strtotime("$regs[1]-$regs[2]-$regs[3] $regs[4]:$regs[5]:$regs[6]Z");
		$result = mktime($regs[4], $regs[5], $regs[6], $regs[2], $regs[3], $regs[1]);
		return $result;
	} else {
		return null;;
	}
	}
	function _normalizeInt($str) {
		$val = intval($str);
		if($val."" === $str)
			return $val;
		else return null;
	}
	function _normalizeDouble($str) {
		$val = doubleval($str);
		if($val."" === $str)
			return $val;
		else return null;
	}
	function _normalizeBoolean($str, $explicit=false) {
		if($explicit)
			if($str == "1") 	return true;
			elseif($str == "0") return false;
			
		if($str=="True")		return true;
		elseif($str=="False")	return false;
		else 				return null;
	}		
	function _normalizeKey($key) {
		$newKey = str_replace("ows_","",$key );
		return ereg_replace("_x[0-9a-f]{4}_","",$newKey);
	}

	//==============================================================================================	
	// DEBUGGING AND ERROR HANDLING
	//==============================================================================================	

	var $debugFunction = null;				// den kaldte funktion
	var $errorcode = null;					// den sidste fejlkode
	var $errorstring = null;				// den sidste fejltekst

	/**
	 * Starter en funktions debugging
	 * @param	string $func Angivelse af hvilken funktion der er blevet kaldt
	 * @access	private
	 */
	function _debugStart($func) {
		$this->result = null;
		$this->client = null;

		$this->errorcode = null;
		$this->errorstring = null;

		if($this->debugFunction==null)
			$this->debugFunction = htmlentities($func);
		else
			$this->debugFuction = "Udefineret sharepoint funktion";
			
		// vis alle funktionskald
		if ($this->debug==2 || $this->debug==5) 
			echo "<hr />Function ".$this->debugFunction;
	}

	/**
	 * Afslutter debugging
	 * @access	private
	 */
	function _debugEnd($rows=null) {
		// vis hvor mange r�kker der blev returneret
		if($this->debug>=2 && $rows!==null)
			echo " Returning {$rows} elements.";

		// vis al kommunikation med serveren
		if ($this->debug>=5) 
			$this->printComm();

		// afslut evt. output
		if ($this->debug==2 || $this->debug>=4) 
			echo "<hr />";

		// slet debug variable
		$this->debugFunction = null;
	}
	/**
	 * Returnerer en opst�et fejl
	 * @access	private
	 */
	function _debugError($errorcode, $description=null) {
		$this->errorcode = $errorcode;
		$this->errorstring = $description;
		if($this->debug>0) { // Der skal vises noget...

			// Hvis der ikke er skrevet noget f�r skriver vi lige hvor det gik galt.
			if ($this->debug !=2 && $this->debug <4)
				echo "<hr />Fejl i ".$this->debugFunction."<br />";

			// Vi skriver en kort fejlmeddelelse
			if ($this->debug>=1) echo $this->getError();

			// Vis al kommunikation med serveren
			if ($this->debug>=3) $this->printComm();
			echo "<hr />";
		}		
		$this->debugFunction = null;
		return false;
	}
	/**
	 * Returner errortype ud fra fejlkkode
	 * @param	int $errorcode fejlkode
	 * @return	string Fejltype
	 */
	function _errortekst($errorcode) {
		switch($errorcode) {
			case 1: return "Construction error";
			case 2: return "Communication error";
			case 3: return "No XML content error";
			case 4: return "Parsing XML error";
			case 5: return "Server error";
			case 10: return "Input error";
		}
	}
	/**
	 * Returnerer fejlkode
	 * @return	int fejlkode
	 * @access	public
	 */
	function getErrorCode() {
		return $this->errorcode;
	}
	/**
	 * Returnerer en beskrivelse af den seneste fejl.
	 * @return	string html beskrivelse af fejl.
	 * @access	public
	 */
	function getError() {
		if ($this->getErrorCode()==null) 
			return null; 
		else
			return "<strong>".$this->getErrorType().":</strong> ".nl2br(htmlentities($this->errorstring));
	}
	/**
	 * Returnerer fejltypen for den seneste fejl.
	 * @return	string fejltype
	 * @access	public
	 */
	function getErrorType() {
		if ($this->getErrorCode()==null)
			return null; 
		else
			return $this->_errortekst($this->getErrorCode());
	}
	/**
	 * Printer alt kommunikation med webservice
	 * @access	private
	 */
	function printComm() {
		echo "<small><p><strong>Request</strong><br /><xmp>".$this->client->request."</xmp></p>".
		     "<p><strong>Response</strong><br /><xmp>". 
		     $this->_getXmlResponse().
		     "</xmp></p></small>";
	}

  } // End of class SharepointSite

  class listQuery{
	var $list   = null; 	// GUID for selve listen
	var $fields = null;	// Hvilke felter der skal vises. Default samtlige felter. Array
	var $where  = null;	// Kriterier. CAML string
	var $order  = null;	// Sorteringsr�kkef�lge. Array
	var $limit  = null;	// Antal r�kker der skal returneres. Default ingen gr�nse. Integer
	var $start  = null;	// Hvis forestp�rgslen skal starte et bestemt sted
	
	function listQuery($list=null) {
		$this->list = $list;
		$this->fields = null;
		$this->limit  = null;
		$this->where  = null;
		$this->order  = null;
	}
	function addViewField($field) {
		if(!is_array($this->fields))
			$this->fields = array();
		$this->fields[] = $field;
	}
	function addOrderField($field, $direction=null) {
		if(!is_array($this->order))
		 	$this->order = array();

		if($direction == null)
			$this->order[] = $field;
		else
			$this->order[$field] = $direction;
	}
	function setWhere($caml) {
		$this->where = $caml;
	}
	function limitRows($limit) {
		$this->limit = intval($limit);
	}
	function startAt($pos) {
		$this->start = intval($pos);
	}
	function _makeBody() {
		$br = "\n";

		$xpath = "/list[@id='".$this->list."']";
		$soapBody = '<queryRequest xmlns="http://schemas.microsoft.com/sharepoint/dsp">'.$br
		          . ' <dsQuery select="'.$xpath.'" resultContent="both" columnMapping="attribute" resultRoot="Rows" resultRow="Row"';
              if($this->start>0) $soapBody .= ' startPosition="'.$this->start.'"';
		$soapBody .= '>'.$br
		          . '  <Query';
		// Rowlimit
		if($this->limit>0) $soapBody .= ' RowLimit="'.$this->limit.'"';
		$soapBody 	.= '>'.$br;
		// Fields
		if(is_array($this->fields) && count($this->fields)>0) {
			$soapBody .= '    <Fields>'.$br;
			foreach($this->fields as $field)
				$soapBody .= '      <Field Name="'.$field.'" />'.$br;
			$soapBody .= '    </Fields>'.$br;

		} else {
			$soapBody .= '    <Fields>'.$br
			           . '      <AllFields IncludeHiddenFields="true" />'.$br
			           . '    </Fields>'.$br;
		}
		// Where
		if($this->where!=null)
			$soapBody .= '    <Where>'.$br. '      '.$this->where.$br. '    </Where>'.$br;
		// Order
		if(is_array($this->order) && count($this->order)>0) {
			$soapBody .= '    <OrderBy>'.$br;
			foreach($this->order as $key => $value) {
				if(strtolower($value) == "desc")
					$soapBody .= '      <OrderField Name="'.$key.'" Direction="DESC" />'.$br;
				elseif(strtolower($value) == "asc")
					$soapBody .= '      <OrderField Name="'.$key.'" Direction="ASC" />'.$br;
				else 
					$soapBody .= '      <OrderField Name="'.$value.'" Direction="ASC" />'.$br;
			}
			$soapBody .= '    </OrderBy>'.$br;
		}


		$soapBody .= '  </Query>'.$br
		           . ' </dsQuery>'.$br
		           . '</queryRequest>';
		return $soapBody;
	}

	// Funktion til at lave en CAML-streng udfra et vilk�rligt antal betingelser
	function makeWhere($handle, $array) {
		$antal = sizeof($array);
		if($antal==0) return "";
		elseif($antal==1) return $array[0];

		$niveauer = ceil( log($antal) / log(2) );
		$caml = array();
		for($i=0; $i<=$niveauer; $i++) $caml[$niveauer] = array();
		foreach($array as $cond) {
			// Saml 2-betingelser til 1 p� h�jere niveau
			if(sizeof($caml[0])>=2) die("Caml overflow");
			// S�rg for at eksisterende betingelser er samlet
			for($i=1; $i<=$niveauer; $i++) {
				if(sizeof($caml[$i])>=2) {
					$caml[($i-1)][] = $this->makeCamlMigrate($handle, $caml[$i]);
					$caml[$i] = array();
				}
			}
			// Inds�t den aktuelle betingelse p� det rette niveau
			for($i=$niveauer; $i>0; $i--) {
				if(sizeof($caml[$niveauer])<2) {
					if(is_array($cond)) {
						$output  = '<'.$cond[0].'><FieldRef Name="'.$cond[1].'" /><Value';
						if($cond[3]) $output .= ' type="'.$cond[3].'"';
						$output .= '>'.$cond[2].'</Value></'.$cond[0].'>';
					} else	$output = $cond;
					$caml[$niveauer][] = $output;
					// Jep det var s� det
					break;
				}
			}
		}
		// Saml stumperne
		// Saml 2-betingerlser til 1 p� h�jere niveau
		for($i=$niveauer; $i>0; $i--) {
			if(sizeof($caml[$i])==2) {
				$caml[$i-1][] = $this->makeCamlMigrate($handle, $caml[$i]);
				$caml[$i] = array();
			}
		}
		// Saml enkelt-st�ende betingelser i den �verste betingelse
		for($p=0; $p<2; $p++)
		for($n=1; $n<=$niveauer; $n++) {
		for($i=$niveauer; $i>0; $i--) {
			if(sizeof($caml[$i])==1) {
				if(!is_array($caml[$i-1])) $caml[$i-1] = array();
				$len = sizeof($caml[$i-1]);
				if($len==1)	$caml[$i-1] = $this->makeCamlMigrate($handle, array($caml[$i-1][0], $caml[$i][0]) );
				elseif($len==0) $caml[$i-1] = $caml[$i];
				else die("Error");
				unset($caml[$i]);
			}
		}
		for($i=$niveauer; $i>0; $i--) {
			if(sizeof($caml[$i])==2) {
				$caml[$i-1][] = $this->makeCamlMigrate($handle, $caml[$i]);
				unset($caml[$i]);
			}
		}
		}

//		if(sizeof($caml[0])==2) return $this->makeCamlMigrate($handle, $
		if(is_array($caml[0]))		return $caml[0][0];
		else	return $caml[0];
	}

	function makeCamlMigrate($handle, $conditions) {
		if($this->debug)
			var_dump($handle, $conditions);
		if(sizeof($conditions)!=2) die("Cannot merge ".sizeof($conditions)." conditions");
		$result = "<{$handle}> "."\r\n     ".$conditions[0]." "."\r\n     ".$conditions[1]."\r\n     "." </{$handle}>";
		if($this->debug)
			var_dump($result);
		return $result;
	}

  } // end of class listQuery

  // Class xml2array
  // Standard komponent
  class xml2Array {
	  
	var $arrOutput;
	var $resParser;
	var $parserResult;
	var $xml;

	function xml2array($doc) {
		$this->xml = $doc;
	}

	function _createParser() {
		$this->resParser = xml_parser_create ();
	       xml_set_object($this->resParser,$this);
		xml_parser_set_option($this->resParser, XML_OPTION_CASE_FOLDING, false);
	       xml_set_element_handler($this->resParser, "tagOpen", "tagClosed");
	       xml_set_character_data_handler($this->resParser, "tagData");
		$this->arrOutput = array();
	}
	function parse() {
		$this->_createParser();
	       $this->parserResult = xml_parse($this->resParser,$this->xml);
		$result = $this->arrOutput;
		if(!$this->parserResult) {
			$errorstring = sprintf("XML error: %s at line %d",
				xml_error_string(xml_get_error_code($this->resParser)),
				xml_get_current_line_number($this->resParser));
			$result = $errorstring;
		}
			xml_parser_free($this->resParser);
		return $result;
	}
	function tagOpen($parser, $name, $attrs) {
		$tag=array("name"=>$name,"attrs"=>$attrs);
		array_push($this->arrOutput,$tag);
	}  
	function tagData($parser, $tagData) {
		if(trim($tagData)) {
			if(isset($this->arrOutput[count($this->arrOutput)-1]['tagData'])) {
				$this->arrOutput[count($this->arrOutput)-1]['chardata'] .= $tagData;
			}
			else {
				$this->arrOutput[count($this->arrOutput)-1]['chardata'] = $tagData;
			}
		}
	}
	function tagClosed($parser, $name) {
		$this->arrOutput[count($this->arrOutput)-2]['children'][] = $this->arrOutput[count($this->arrOutput)-1];
		array_pop($this->arrOutput);
	}
  } // end of class xml2array

} // Defined
?>