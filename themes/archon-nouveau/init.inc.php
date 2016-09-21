<?php
$_ARCHON->PublicInterface->Delimiter = " &rarr; ";


function list_item($item, $attributes = false){

	$li = "\t<li";

	if(!empty($attributes)){
		foreach($attributes as $name => $value){
			$li .= " $name=\"$value\"";
		}
	}
	$li .= ">$item</li>\n";

	return $li;
}

?>