<?php
/**
 * Output file for browsing by subject
 *
 * @package Archon
 * @author Chris Rishel
 */

isset($_ARCHON) or die();



$in_ID = isset($_REQUEST['id']) ? $_REQUEST['id'] : NULL;
$in_Char = isset($_REQUEST['char']) ? $_REQUEST['char'] : NULL;
$in_SubjectTypeID = isset($_REQUEST['subjecttypeid']) ? $_REQUEST['subjecttypeid'] : 0;
$in_Browse = isset($_REQUEST['browse']) ? true : false;



$objSubjectsTitlePhrase = Phrase::getPhrase('subjects_title', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
$strSubjectsTitle = $objSubjectsTitlePhrase ? $objSubjectsTitlePhrase->getPhraseValue(ENCODE_HTML) : 'Browse by Subject';

$_ARCHON->PublicInterface->Title = $strSubjectsTitle;
$_ARCHON->PublicInterface->addNavigation($_ARCHON->PublicInterface->Title, "?p={$_REQUEST['p']}");
$_ARCHON->PublicInterface->Title .= ' | ' . $_ARCHON->Repository->Name;

if($in_SubjectTypeID)
{
   $objSubjectType = New SubjectType($in_SubjectTypeID);
   $objSubjectType->dbLoad();
   $_ARCHON->PublicInterface->addNavigation(str_replace('$1', $objSubjectType->toString(), $strOfType), "?p={$_REQUEST['p']}&amp;subjecttypeid=$in_SubjectTypeID");
}

if($in_ID)
{
   $vars = subjects_listChildSubjects($in_ID);
}
elseif($in_Char)
{
   $vars = subjects_listSubjectsForChar($in_Char, $in_SubjectTypeID);
}
elseif($in_Browse)
{
   $in_Page = $_REQUEST['page'] ? $_REQUEST['page'] : 1;

   $vars = subjects_listAllSubjects($in_Page, $in_SubjectTypeID);
}
else
{
   $vars = subjects_main($in_SubjectTypeID);
}

require_once("header.inc.php");

echo($_ARCHON->PublicInterface->executeTemplate('subjects', 'SubjectNav', $vars));

require_once("footer.inc.php");


function subjects_main($SubjectTypeID)
{
   global $_ARCHON;

   $objOfTypePhrase = Phrase::getPhrase('subjects_oftype', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
   $strOfType = $objOfTypePhrase ? $objOfTypePhrase->getPhraseValue(ENCODE_HTML) : 'Of Type "$1"';


   $objViewAllPhrase = Phrase::getPhrase('viewall', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
   $strViewAll = $objViewAllPhrase ? $objViewAllPhrase->getPhraseValue(ENCODE_HTML) : 'View All';


   $arrSubjectCount = $_ARCHON->countSubjects(true, $SubjectTypeID);
   $arrSubjectTypes = $_ARCHON->getAllSubjectTypes();

   $vars['strPageTitle'] = strip_tags($_ARCHON->PublicInterface->Title);
   $vars['strSubTitleClasses'] = 'bold';
   $vars['strBackgroundID'] = '';

	$strShowBeginning = '';

   if($arrSubjectTypes[$SubjectTypeID])
   {
      $objTypedShowBeginningWithPhrase = Phrase::getPhrase('subjects_typedshowbeginningwith', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
      $strTypedShowBeginningWith = $objTypedShowBeginningWithPhrase ? $objTypedShowBeginningWithPhrase->getPhraseValue(ENCODE_HTML) : 'Show "$1" Subjects Beginning with';
      $strTypedShowBeginningWith = str_replace('$1', $arrSubjectTypes[$SubjectTypeID]->toString(), $strTypedShowBeginningWith);

      $strShowBeginning = $strTypedShowBeginningWith;
   }
   else
   {
      $objShowBeginningWithPhrase = Phrase::getPhrase('subjects_showbeginningwith', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
      $strShowBeginningWith = $objShowBeginningWithPhrase ? $objShowBeginningWithPhrase->getPhraseValue(ENCODE_HTML) : 'Show Subjects Beginning with';

      $strShowBeginning = $strShowBeginningWith;
   }

   $vars['strSubTitle'] = $strShowBeginning.":";

   $content = "<div class=\"center\">\n";

   if(!empty($arrSubjectCount['#']))
   {
      $href = "?p={$_REQUEST['p']}&amp;char=" . urlencode('#');
      if($SubjectTypeID)
      {
         $href .= "&amp;subjecttypeid=$SubjectTypeID";
      }
      $content .= "<a href='$href'>-#-</a>" . INDENT;
   }
   else
   {
      $content .= "-#-" . INDENT;
   }

   for($i = 65; $i < 91; $i++)
   {
      $char = chr($i);

      if(!empty($arrSubjectCount[encoding_strtolower($char)]))
      {
         $href = "?p={$_REQUEST['p']}&amp;char=$char";
         if($SubjectTypeID)
         {
            $href .= "&amp;subjecttypeid=$SubjectTypeID";
         }
         $content .= "<a href='$href'>-$char-</a>" . INDENT;
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
   $content .= "<br /><br /><a href='?p={$_REQUEST['p']}&amp;browse&amp;subjecttypeid={$SubjectTypeID}'>{$strViewAll}</a>";
   $content .= "</div>\n";

   $objFilterByPhrase = Phrase::getPhrase('subjects_filterby', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
   $strFilterBy = $objFilterByPhrase ? $objFilterByPhrase->getPhraseValue(ENCODE_HTML) : 'Filter Subjects by';

	$vars['strFilterBy'] = $strFilterBy;

   $subTopics = '';

   if(!empty($arrSubjectTypes))
   {
      foreach($arrSubjectTypes as $objSubjectType)
      {
         if($objSubjectType->ID != $SubjectTypeID)
         {
            $subTopics .= "<a href='?p={$_REQUEST['p']}&amp;subjecttypeid=$objSubjectType->ID'>" . $objSubjectType->toString() . "</a><br />";
         }
         else
         {
            $subTopics .= "{$objSubjectType->toString()}<br />";
         }
      }
   }



   $vars['content'] = $content;
   $vars['subTopics'] = $subTopics;

   return $vars;
}


function subjects_listChildSubjects($ID)
{
   global $_ARCHON;

   $objSubject = New Subject($ID);
   $objSubject->dbLoad();

   $objSubTermHeaderPhrase = Phrase::getPhrase('subjects_subtermheader', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
   $strSubTermHeader = $objSubTermHeaderPhrase ? $objSubTermHeaderPhrase->getPhraseValue(ENCODE_HTML) : 'Sub-Terms Under $1';
   $strSubTermHeader = str_replace('$1', $objSubject->toString(LINK_NONE, true, $_ARCHON->PublicInterface->Delimiter), $strSubTermHeader);
   $objRelatedRecordsPhrase = Phrase::getPhrase('subjects_relatedrecords', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
   $strRelatedRecords = $objRelatedRecordsPhrase ? $objRelatedRecordsPhrase->getPhraseValue(ENCODE_HTML) : 'Related Records';


   $_ARCHON->PublicInterface->addNavigation($objSubject->toString(LINK_EACH, true, $_ARCHON->PublicInterface->Delimiter), "?p={$_REQUEST['p']}&amp;id=$ID");
   $arrSubjects = $_ARCHON->getChildSubjects($ID);


   if(empty($arrSubjects))
   {
      header("Location: index.php?p=core/search&subjectid=$ID");
   }else
   {

      if(!$_ARCHON->PublicInterface->Templates[$_ARCHON->Package->APRCode]['SubjectList'])
      {
         $_ARCHON->declareError("Could not list Subjects: SubjectList template not defined for template set {$_ARCHON->PublicInterface->TemplateSet}.");
      }

      $vars['strPageTitle'] = strip_tags($_ARCHON->PublicInterface->Title);
      $vars['strSubTitleClasses'] = 'listitemhead bold';
      $vars['strSubTitle'] = $strSubTermHeader.':';
      $vars['strBackgroundID'] = 'listitemwrapper';

      $content = '';

      $content .= "<span class='small' style='margin-left:.5em'>(<a href='?p=core/search&amp;subjectid=$ID'>$strRelatedRecords</a>)</span>";


      foreach($arrSubjects as $objSubject)
      {
		ob_start();
        eval($_ARCHON->PublicInterface->Templates[$_ARCHON->Package->APRCode]['SubjectList']);
		$content .= ob_get_contents();
		ob_end_clean();
      }
   }

   $vars['content'] = $content;
   $vars['subTopics'] = false;

   return $vars;
}


function subjects_listSubjectsForChar($Char, $SubjectTypeID)
{
   global $_ARCHON;

   $objTypedBeginningWithPhrase = Phrase::getPhrase('subjects_typedbeginningwith', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
   $strTypedBeginningWith = $objTypedBeginningWithPhrase ? $objTypedBeginningWithPhrase->getPhraseValue(ENCODE_HTML) : 'Of Type "$1" Beginning with "$2"';
   $objSubjectsBeginningWithPhrase = Phrase::getPhrase('subjects_subjectsbeginningwith', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
   $strSubjectsBeginningWith = $objSubjectsBeginningWithPhrase ? $objSubjectsBeginningWithPhrase->getPhraseValue(ENCODE_HTML) : 'Beginning with "$1"';


   $arrSubjectTypes = $_ARCHON->getAllSubjectTypes();


   if($SubjectTypeID)
   {
      $_ARCHON->PublicInterface->addNavigation(str_replace(array('$1', '$2'), array($arrSubjectTypes[$SubjectTypeID]->toString(), encoding_strtoupper($Char)), $strTypedBeginningWith), "?p={$_REQUEST['p']}&amp;subjecttypeid=$SubjectTypeID&amp;char=$Char");
   }
   else
   {
      $_ARCHON->PublicInterface->addNavigation(str_replace('$1', encoding_strtoupper($Char), $strSubjectsBeginningWith), "?p={$_REQUEST['p']}&amp;char=$Char");
   }


   if(!$_ARCHON->PublicInterface->Templates[$_ARCHON->Package->APRCode]['SubjectList'])
   {
      $_ARCHON->declareError("Could not list Subjects: SubjectList template not defined for template set {$_ARCHON->PublicInterface->TemplateSet}.");
   }


   $vars['strPageTitle'] = strip_tags($_ARCHON->PublicInterface->Title);
   $vars['strSubTitleClasses'] = 'listitemhead bold';
   $vars['strBackgroundID'] = '';

   $content = '';

   $arrSubjects = $_ARCHON->getSubjectsForChar($Char, $SubjectTypeID);

   if(!empty($arrSubjects))
   {
      if($arrSubjectTypes[$SubjectTypeID])
      {
         $objTypedBeginningWithHeaderPhrase = Phrase::getPhrase('subjects_typedbeginningwithheader', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
         $strTypedBeginningWithHeader = $objTypedBeginningWithHeaderPhrase ? $objTypedBeginningWithHeaderPhrase->getPhraseValue(ENCODE_HTML) : '"$1" Subjects Beginning with "$2"';
         $strTypedBeginningWithHeader = str_replace(array('$1', '$2'), array($arrSubjectTypes[$SubjectTypeID]->toString(), encoding_strtoupper($Char)), $strTypedBeginningWithHeader);

         $vars['strSubTitle'].= $strTypedBeginningWithHeader;
      }
      else
      {
         $objSubjectsBeginningWithHeaderPhrase = Phrase::getPhrase('subjects_subjectsbeginningwithheader', PACKAGE_SUBJECTS, 0, PHRASETYPE_PUBLIC);
         $strSubjectsBeginningWithHeader = $objSubjectsBeginningWithHeaderPhrase ? $objSubjectsBeginningWithHeaderPhrase->getPhraseValue(ENCODE_HTML) : 'Subjects Beginning with "$1"';
         $strSubjectsBeginningWithHeader = str_replace('$1', encoding_strtoupper($Char), $strSubjectsBeginningWithHeader);

         $vars['strSubTitle'] = $strSubjectsBeginningWithHeader;
      }

      foreach($arrSubjects as $objSubject)
      {
         ob_start();
         eval($_ARCHON->PublicInterface->Templates[$_ARCHON->Package->APRCode]['SubjectList']);
         $content .= ob_get_contents();
         ob_end_clean();
      }
   }

   $vars['content'] = $content;
   $vars['subTopics'] = false;

   return $vars;
}


function subjects_listAllSubjects($Page, $SubjectTypeID)
{
   global $_ARCHON;

   $arrSubjects = $_ARCHON->searchSubjects($_REQUEST['q'], NULL, $SubjectTypeID, CONFIG_CORE_PAGINATION_LIMIT + 1, ($Page-1)*CONFIG_CORE_PAGINATION_LIMIT);

   if(count($arrSubjects) > CONFIG_CORE_PAGINATION_LIMIT)
   {
      $morePages = true;
      array_pop($arrSubjects);
   }

// Set up a URL for any prev/next buttons or in case $Page
// is too high
   $paginationURL = 'index.php?p=' . $_REQUEST['p'].'&browse&subjecttypeid='.$SubjectTypeID;

   if(empty($arrSubjects) && $Page != 1)
   {
      header("Location: $paginationURL");
   }

   

   $objViewAllPhrase = Phrase::getPhrase('viewall', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
   $strViewAll = $objViewAllPhrase ? $objViewAllPhrase->getPhraseValue(ENCODE_HTML) : 'View All';

   $_ARCHON->PublicInterface->addNavigation($strViewAll);


   if(!$_ARCHON->PublicInterface->Templates[$_ARCHON->Package->APRCode]['SubjectList'])
   {
      $_ARCHON->declareError("Could not list Subjects: SubjectList template not defined for template set {$_ARCHON->PublicInterface->TemplateSet}.");
   }

   $vars['strPageTitle'] = strip_tags($_ARCHON->PublicInterface->Title);
   $vars['strSubTitleClasses'] = 'listitemhead bold';
   $vars['strBackgroundID'] = '';

   $content = '';

   if(!$_ARCHON->Error)
   {
      if(!empty($arrSubjects))
      {
         $vars['strSubTitle'] = $strViewAll;

         foreach($arrSubjects as $objSubject)
         {
            ob_start();
            eval($_ARCHON->PublicInterface->Templates[$_ARCHON->Package->APRCode]['SubjectList']);
            $content .= ob_get_contents();
            ob_end_clean();
         }
      }


      if($Page > 1 || $morePages)
      {
         $pages = '';
         $pages .= "<div class='paginationnav'>";

         if($Page > 1)
         {
            $prevPage = $Page - 1;
            $prevURL = encode($paginationURL . "&page=$prevPage", ENCODE_HTML);
            $pages .= "<span class='paginationprevlink'><a href='$prevURL'>Prev</a></span>";
         }
         if($morePages)
         {
            $nextPage = $Page + 1;
            $nextURL = encode($paginationURL . "&page=$nextPage", ENCODE_HTML);
            $pages .= "<span class='paginationnextlink'><a href='$nextURL'>Next</a></span>";
         }
         $pages .= "</div>";
      }
   }

   $vars['content'] = $content;
   if(isset($pages)){ $vars['pages'] = $pages; }

   return $vars;
}





?>
