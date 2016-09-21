<?php
/**
 * Output file for browsing by creator
 *
 * @package Archon
 * @author Chris Rishel
 */
isset($_ARCHON) or die();

if(!$_ARCHON->PublicInterface->Templates['creators']['Creator'])
{
   $_ARCHON->declareError("Could not display Creator: Creator template not defined for template set {$_ARCHON->PublicInterface->TemplateSet}.");
}


$in_Char = isset($_REQUEST['char']) ? $_REQUEST['char'] : NULL;

$in_Browse = isset($_REQUEST['browse']) ? true : false;


$objCreatorsTitlePhrase = Phrase::getPhrase('creators_title', PACKAGE_CREATORS, 0, PHRASETYPE_PUBLIC);
$strCreatorsTitle = $objCreatorsTitlePhrase ? $objCreatorsTitlePhrase->getPhraseValue(ENCODE_HTML) : 'Browse by Creator';
$_ARCHON->PublicInterface->Title = $strCreatorsTitle;
$_ARCHON->PublicInterface->addNavigation($_ARCHON->PublicInterface->Title, "?p={$_REQUEST['p']}");
$_ARCHON->PublicInterface->Title .= ' | '.$_ARCHON->Repository->Name;


if($in_Char)
{
   $vars = creators_listCreatorsForChar($in_Char);
}
elseif($in_Browse)
{
   $in_Page = $_REQUEST['page'] ? $_REQUEST['page'] : 1;

   $vars = creators_listAllCreators($in_Page);
}
else
{
   $vars = creators_main();
}

require_once("header.inc.php");

echo($_ARCHON->PublicInterface->executeTemplate('creators', 'CreatorNav', $vars));

require_once("footer.inc.php");


function creators_main()
{
   global $_ARCHON;

   

   $objShowBeginningWithPhrase = Phrase::getPhrase('creators_showbeginningwith', PACKAGE_CREATORS, 0, PHRASETYPE_PUBLIC);
   $strShowBeginningWith = $objShowBeginningWithPhrase ? $objShowBeginningWithPhrase->getPhraseValue(ENCODE_HTML) : 'Show Creators Beginning with';

   $objViewAllPhrase = Phrase::getPhrase('viewall', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
   $strViewAll = $objViewAllPhrase ? $objViewAllPhrase->getPhraseValue(ENCODE_HTML) : 'View All';

   $arrCreatorCount = $_ARCHON->countCreators(true);

	$vars['strPageTitle'] = strip_tags($_ARCHON->PublicInterface->Title);
	$vars['strSubTitle'] = $strShowBeginningWith.":";
	$vars['strSubTitleClasses'] = 'bold center';
	$vars['strBackgroundID'] = '';

	$content = "<div class=\"center\">\n";

	 if(!empty($arrCreatorCount['#']))
	 {
		$content .= "<a href='?p={$_REQUEST['p']}&amp;char=" . urlencode('#'). "'>-#-</a>" . INDENT;
	 }
	 else
	 {
		$content .= "-#-" . INDENT;
	 }

	 for($i = 65; $i < 91; $i++)
	 {
		$char = chr($i);

		if(!empty($arrCreatorCount[encoding_strtolower($char)]))
		{
		   $content .= "<a href='?p={$_REQUEST['p']}&amp;char=$char'>-$char-</a>" . INDENT;
		}
		else
		{
		   $content .= "-$char-" . INDENT;
		}

		if($char == 'M')
		{
		   $content .= "<br /><br />\n";
		}
	 }

	$content .= "<br /><br /><a href='?p={$_REQUEST['p']}&amp;browse'>{$strViewAll}</a>";
	$content .= "</div>\n";

	$vars['content'] = $content;

	return $vars;
}


function creators_listCreatorsForChar($Char)
{
   global $_ARCHON;

   

   $objNavBeginningWithPhrase = Phrase::getPhrase('creators_navbeginningwith', PACKAGE_CREATORS, 0, PHRASETYPE_PUBLIC);
   $strNavBeginningWith = $objNavBeginningWithPhrase ? $objNavBeginningWithPhrase->getPhraseValue(ENCODE_HTML) : 'Beginning with "$1"';
   $objCreatorsBeginningWithPhrase = Phrase::getPhrase('creators_creatorsbeginningwith', PACKAGE_CREATORS, 0, PHRASETYPE_PUBLIC);
   $strCreatorsBeginningWith = $objCreatorsBeginningWithPhrase ? $objCreatorsBeginningWithPhrase->getPhraseValue(ENCODE_HTML) : 'Creators Beginning with "$1"';

   $_ARCHON->PublicInterface->addNavigation(str_replace('$1', encoding_strtoupper($Char), $strNavBeginningWith), "?p={$_REQUEST['p']}&amp;char=$Char");

   if(!$_ARCHON->PublicInterface->Templates[$_ARCHON->Package->APRCode]['CreatorList'])
   {
      $_ARCHON->declareError("Could not list Creators: CreatorList template not defined for template set {$_ARCHON->PublicInterface->TemplateSet}.");
   }

   $vars['strPageTitle'] = strip_tags($_ARCHON->PublicInterface->Title);
   $vars['strSubTitle'] = str_replace('$1', encoding_strtoupper($Char), $strCreatorsBeginningWith);
   $vars['strSubTitleClasses'] = 'listitemhead bold';
   $vars['strBackgroundID'] = ' id="listitemwrapper"';
   $content = '';

   if(!$_ARCHON->Error)
   {
      $arrCreators = $_ARCHON->getCreatorsForChar($Char);

      if(!empty($arrCreators))
      {
         foreach($arrCreators as $objCreator)
         {
            $content .= $_ARCHON->PublicInterface->executeTemplate('creators', 'CreatorList', array('objCreator' => $objCreator));
         }
      }
   }

   $vars['content'] = $content;
   return $vars;
}

function creators_listAllCreators($Page)
{
   global $_ARCHON;

   $arrCreators = $_ARCHON->searchCreators($_REQUEST['q'], CONFIG_CORE_PAGINATION_LIMIT + 1, ($Page-1)*CONFIG_CORE_PAGINATION_LIMIT);

   if(count($arrCreators) > CONFIG_CORE_PAGINATION_LIMIT)
   {
      $morePages = true;
      array_pop($arrCreators);
   }

// Set up a URL for any prev/next buttons or in case $Page
// is too high
   $paginationURL = 'index.php?p=' . $_REQUEST['p'].'&browse';

   if(empty($arrCreators) && $Page != 1)
   {
      header("Location: $paginationURL");
   }

   

   $objViewAllPhrase = Phrase::getPhrase('viewall', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
   $strViewAll = $objViewAllPhrase ? $objViewAllPhrase->getPhraseValue(ENCODE_HTML) : 'View All';

   $_ARCHON->PublicInterface->addNavigation($strViewAll);


   if(!$_ARCHON->PublicInterface->Templates[$_ARCHON->Package->APRCode]['CreatorList'])
   {
      $_ARCHON->declareError("Could not list Creators: CreatorList template not defined for template set {$_ARCHON->PublicInterface->TemplateSet}.");
   }

   $vars['strPageTitle'] = strip_tags($_ARCHON->PublicInterface->Title);
   $vars['strSubTitle'] = $strViewAll;
   $vars['strSubTitleClasses'] = 'listitemhead bold';
   $vars['strBackgroundID'] = '';
   $content = '';

   if(!$_ARCHON->Error)
   {
      if(!empty($arrCreators))
      {
         foreach($arrCreators as $objCreator)
         {
			$content .= $_ARCHON->PublicInterface->executeTemplate('creators', 'CreatorList', array('objCreator' => $objCreator));
         }
      }

      if($Page > 1 || $morePages)
      {
         $content .= "<div class='paginationnav'>";

         if($Page > 1)
         {
            $prevPage = $Page - 1;
            $prevURL = encode($paginationURL . "&page=$prevPage", ENCODE_HTML);
            $content .= "<span class='paginationprevlink'><a href='$prevURL'>Prev</a></span>";
         }
         if($morePages)
         {
            $nextPage = $Page + 1;
            $nextURL = encode($paginationURL . "&page=$nextPage", ENCODE_HTML);
            $content .= "<span class='paginationnextlink'><a href='$nextURL'>Next</a></span>";
         }
         $content .= "</div>";
      }
   }

	$vars['content'] = $content;
	return $vars;
}
