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

$is_index = $strSubTitle == "Show Digital Content Titles Beginning with:"? true : false;
$links = array();


if($is_index){
	$content = str_replace('</div>', '', $content);
	$content = str_replace('<div class="center">', '', $content);
	$content = str_replace('<br />', '', $content);
	$content = str_replace('<br/>', '', $content);
	$content = str_replace("\n", '', $content);
	$content = str_replace("\r", '', $content);

	$links = explode('&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;', $content);

	$all = array_pop($links);

	for($i = 0; $i < count($links); $i++){
		$links[$i] = str_replace('-', '', $links[$i]);
		$links[$i] = str_replace('archonnouveau', 'archon-nouveau', $links[$i]);
	}

	$last_parts = explode('</a>      <form', $all);
	$last_parts[0] .= '</a>';
	$all = $last_parts[0];

	$image_search = '<form'.$last_parts[1];
}

?>

<h2><?php print $heading; ?></h2>


<?php if($is_index){ ?>

<div class="results-list">
	<ul>
		<?php
				foreach($links as $link){ print "\t<li$class>$link</li>\n"; }
				print "\t\t<li class=\"view-all\">$all</li>\n";
		?>
	</ul>
</div>

<?php print $image_search; ?>

<?php if(isset($pages)){ echo($pages); } ?>


<?php } else { print $content; } ?>
