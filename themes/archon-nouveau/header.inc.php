<?php
/**
 * Header file for Archon Nouveau theme
 *
 * @package Archon
 * @author Will Martin
 */
isset($_ARCHON) or die();

if($_ARCHON->Script == 'packages/collections/pub/findingaid.php')
{
   require("faheader.inc.php");
   return;
}

$_ARCHON->PublicInterface->Header->OnLoad .= "";

if($_ARCHON->Error)
{
   $_ARCHON->PublicInterface->Header->OnLoad .= " alert('" . encode(str_replace(';', "\n", $_ARCHON->processPhrase($_ARCHON->Error)), ENCODE_JAVASCRIPT) . "');";
}

if(defined('PACKAGE_COLLECTIONS'))
{

   if($objCollection->Repository)
   {
      $RepositoryName = $objCollection->Repository->getString('Name');
   }
   elseif($objDigitalContent->Collection->Repository)
   {
      $RepositoryName = $objDigitalContent->Collection->Repository->getString('Name');
   }
   else
   {
      $RepositoryName = $_ARCHON->Repository ? $_ARCHON->Repository->getString('Name') : '';
   }

   $_ARCHON->PublicInterface->Title = $_ARCHON->PublicInterface->Title ? $_ARCHON->PublicInterface->Title . ' | ' . $RepositoryName : $RepositoryName;

   if($_ARCHON->QueryString && $_ARCHON->Script == 'packages/core/pub/search.php')
   {
      $_ARCHON->PublicInterface->addNavigation("Search Results For \"" . $_ARCHON->getString(QueryString) . "\"", "?p=core/search&amp;q=" . $_ARCHON->QueryStringURL, true);
   }
}
else
{
   $RepositoryName = $_ARCHON->Repository ? $_ARCHON->Repository->getString('Name') : 'Archon';

   $_ARCHON->PublicInterface->Title = $_ARCHON->PublicInterface->Title ? $_ARCHON->PublicInterface->Title . ' | ' . $RepositoryName : $RepositoryName;

   if($_ARCHON->QueryString)
   {
      $_ARCHON->PublicInterface->addNavigation("Search Results For \"" . encode($_ARCHON->QueryString, ENCODE_HTML) . "\"", "?p=core/search&amp;q=" . $_ARCHON->QueryStringURL, true);
   }
}

$_ARCHON->PublicInterface->addNavigation('Archon', 'index.php', true);


///////////////////////////////////////////////////////////////////////////////////////////////


// Set 'active' status on browse links based on the URL.
$arrP = explode('/', $_REQUEST['p']);
$TitleClass				= ($arrP[0] == 'collections' && $arrP[1] != 'classifications')? 'active' : '';
$ClassificationsClass	= ($arrP[1] == 'classifications')? 'active' : '';
$SubjectsClass			= ($arrP[0] == 'subjects')? 'active' : '';
$CreatorsClass			= ($arrP[0] == 'creators')? 'active' : '';
$DigitalLibraryClass	= ($arrP[0] == 'digitallibrary')? 'active' : '';

// Set up the login/welcome/logout/account/cart links

$links = array();

// Login/Welcome/Logout
if($_ARCHON->Security->isAuthenticated()){

	$name_parts = explode(' (', $_ARCHON->Security->Session->User->toString());
	$name = $name_parts[0];

	$logoutURI = preg_replace('/(&|\\?)f=([\\w])*/', '', $_SERVER['REQUEST_URI']);
	$Logout = (encoding_strpos($logoutURI, '?') !== false) ? '&amp;f=logout' : '?f=logout';
	$strLogout = encode($logoutURI, ENCODE_HTML) . $Logout;

	$links['welcome first'] = "Welcome, <span class=\"name\">$name</span> (<a href=\"$strLogout\">Logout</a>)";

} elseif($_ARCHON->config->ForceHTTPS) {

	$links['login first'] = '<a href="index.php?p=core/login&amp;go=">Log In</a>';

} else {
	$links['login first'] = '<a href="#" id="logintrigger">Log In</a>';
}

if($_ARCHON->Security->userHasAdministrativeAccess()){
	$links['admin'] = "<a href='?p=admin' rel='external'>Admin</a>&nbsp;";
}

$referrer = urlencode($_SERVER['HTTP_HOST'] . $_SERVER['REQUEST_URI']);

// Contact
if(!$_ARCHON->Security->userHasAdministrativeAccess()) {

	$emailpage = defined('PACKAGE_COLLECTIONS') ? "collections/research" : "core/contact";

	$links['contact'] = "<a href=\"?p=$emailpage&amp;f=email&amp;referer=$referrer\">Contact Us</a>";

	if($_ARCHON->Security->isAuthenticated()){
		$links['account'] = '<a href=\"?p=core/account&amp;f=account\">My Account</a>';
	}

	if(defined('PACKAGE_COLLECTIONS')){
		$_ARCHON->Security->Session->ResearchCart->getCart();

		$EntryCount = $_ARCHON->Security->Session->ResearchCart->getCartCount();
		//$class = $_ARCHON->Repository->ResearchFunctionality & RESEARCH_COLLECTIONS ? '' : 'hidewhenempty';
		$class = 'cart';
		$hidden = ($_ARCHON->Repository->ResearchFunctionality & RESEARCH_COLLECTIONS || $EntryCount) ? false : true;
		if($hidden){ $class .= ' no-show'; }

		$links[$class] = "<a href=\"?p=collections/research&amp;f=cart&amp;referer=$referrer\">View Cart (<span id=\"cartcount\">$EntryCount</span>)</a>";
	}
}


// Set messages, if any.
$message = false;
if($_ARCHON->PublicInterface->Header->Message && $_ARCHON->PublicInterface->Header->Message != $_ARCHON->Error)
{
	$message = '<div class="message">' . encode($_ARCHON->PublicInterface->Header->Message, ENCODE_HTML) . "</div>\n";
}

// Not entirely sure what this does, but it sets an option in the search box.
$search_list = false;
if(defined('PACKAGE_COLLECTIONS') && CONFIG_COLLECTIONS_SEARCH_BOX_LISTS)
{
	$search_list = '<input type="hidden" name="content" value="1" />';
}


$page_title = strip_tags($_ARCHON->PublicInterface->Title);

$theme_path = $_ARCHON->PublicInterface->Theme;
$image_path = $_ARCHON->PublicInterface->ImagePath;
$js_path = $_ARCHON->PublicInterface->ThemeJavascriptPath;

$breadcrumbs = $_ARCHON->PublicInterface->createNavigation();

$onLoad = $_ARCHON->PublicInterface->Header->OnLoad;
$onUnload = $_ARCHON->PublicInterface->Header->OnUnload;

$search_value = encode($_ARCHON->QueryString, ENCODE_HTML);

?><!DOCTYPE html>
<html lang="en">
<head>
	<meta charset="UTF-8" />
	<title><?php print $page_title; ?></title>

	<link rel="shortcut icon" href="<?php print $image_path; ?>/favicon.ico"/>

	<link rel="stylesheet" type="text/css" href="themes/<?php print $theme_path; ?>/style.css" media="screen" />
	<link rel="stylesheet" type="text/css" href="themes/<?php print $theme_path; ?>/colorbox.css" media="screen" />
	<link rel="stylesheet" type="text/css" href="themes/<?php print $theme_path; ?>/print.css" media="print" />

	<!--[if lte IE 8]>
		<link rel="stylesheet" type="text/css" href="themes/<?php print $theme_path; ?>/ie.css" />
	<![endif]-->

	<script type="text/javascript" src="<?php print $js_path; ?>/jquery-1.7.2.min.js"></script>
	<script type="text/javascript" src="<?php print $js_path; ?>/jquery.colorbox.min.js"></script>

	<?php print $_ARCHON->getJavascriptTags('archon'); ?>

	<script type="text/javascript">
		/* <![CDATA[ */
		var imagePath = '<?php print $image_path; ?>';

		$(document).ready(function(){
			$('#logintrigger').colorbox({width:'380px', height: '270px', inline: true, href: '#userlogin'});
		});

		$(document).ready(function() {<?php print $onLoad; ?>});
		$(window).unload(function() {<?php print $onUnload; ?>});
		/* ]]> */
	</script>
</head>
<body>

<?php if($message){ print $message; } ?>
<a href="#main" class="visually-hidden">Skip to main content</a>
<div id="frame">
<header>
<ul id="utility-nav">
<?php
foreach($links as $class => $link){ print "\t<li class=\"$class\">$link</li>\n"; }
?>
</ul>
<a href="./"><img src="<?php print $image_path; ?>/archon-logo.png" alt="Archon" width="322" height="70" /></a>
	<nav>
		<ul>
			<li class="first">
				<form action="index.php" accept-charset="UTF-8" method="get" id="archon-search-form">
				<label for="archon-search" class="visually-hidden">Search Finding Aids</label>
				<input
					type="search"
					size="25"
					title="search"
					maxlength="150"
					name="q"
					id="archon-search"
					class="search-box"
					placeholder="Search Finding Aids"
					value="<?php print $search_value; ?>"
				/>

				<input type="hidden" name="p" value="core/search" />
				<?php if($search_list){ print $search_list; } ?>
				
				<input type="submit" value="Go" />
				</form>
			</li>
			<li class="<?php print $TitleClass; ?>"><a href="?p=collections/collections">Collections</a></li>
			<li class="<?php print $DigitalLibraryClass; ?>"><a href="?p=digitallibrary/digitallibrary">Digital Content</a></li>
			<li class="<?php print $SubjectsClass; ?>"><a href="?p=subjects/subjects">Subjects</a></li>
			<li class="<?php print $CreatorsClass; ?>"><a href="?p=creators/creators">Creators</a></li>
			<li class="<?php print $ClassificationsClass; ?> last"><a href="?p=collections/classifications">Record Groups</a></li>
		</ul>
	</nav>
</header>

<section>
<div id="breadcrumbs"><?php print $breadcrumbs; ?></div>
<a name="main" id="main"></a>