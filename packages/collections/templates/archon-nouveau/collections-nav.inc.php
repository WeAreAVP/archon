<?php
/**
 * Digital content navigation template
 *
 *
 * The Archon API is also available through the variable:
 *
 *  $_ARCHON
 *
 * Refer to the Archon class definition in lib/archon.inc.php
 * for available properties and methods.
 *
 * @package Archon
 * @author Will Martin
 */

isset($_ARCHON) or die();
$parts = explode(' | ', $strPageTitle);
$heading = $parts[0];

$is_index = $strSubTitle == "Browse Holdings Beginning With:"? true : false;
$links = array();


if($is_index){
	$content = str_replace('</div>', '', $content);
	$content = str_replace('<div class="center">', '', $content);
	$content = str_replace('<br />', '', $content);
	$content = str_replace('<br/>', '', $content);
	$content = str_replace("\n", '', $content);
	$content = str_replace("\r", '', $content);

	$links = explode('&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;', $content);

	for($i = 0; $i < count($links); $i++){
		$links[$i] = str_replace('-', '', $links[$i]);
		$links[$i] = str_replace('archonnouveau', 'archon-nouveau', $links[$i]);
	}
}

?>

<h2><?php print $heading; ?></h2>

<?php if($is_index){ ?>


<div class="results-list">
	<ul>
		<?php
			$all = array_pop($links);
				foreach($links as $link){ print "\t<li$class>$link</li>\n"; }
				print "\t\t<li class=\"view-all\">$all</li>\n";
		?>
	</ul>
</div>

<?php } else { print $content; } ?>

<?php if(isset($pages)){ print $pages; } ?>
