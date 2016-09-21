<?php
/**
 * Public user contact form
 *
 * @package Archon
 * @author Paul Sorensen
 */

isset($_ARCHON) or die();

if($_ARCHON->Security->isAuthenticated() && $_ARCHON->Security->userHasAdministrativeAccess())
{
    header('Location: index.php?p=');
}


contact_initialize();

function contact_initialize()
{
	if(!isset($_REQUEST['f']))
	{
		$_REQUEST['f'] = 'email';
	}

    if($_REQUEST['f'] == 'email')
    {
        contact_email();
    }
    else
    {
        contact_exec();
    }
}



function contact_email()
{
    global $_ARCHON;

    

    $objEmailTitlePhrase = Phrase::getPhrase('research_email_title', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strEmailTitle = $objEmailTitlePhrase ? $objEmailTitlePhrase->getPhraseValue(ENCODE_HTML) : 'Send Email';
    $objExtPhrase = Phrase::getPhrase('research_email_ext', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strExt = $objExtPhrase ? $objExtPhrase->getPhraseValue(ENCODE_HTML) : 'Ext';
    $objMarkedPhrase = Phrase::getPhrase('research_email_marked', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strMarked = $objMarkedPhrase ? $objMarkedPhrase->getPhraseValue(ENCODE_NONE) : 'Fields marked with an asterisk (<span style="color:red;">*</span>) are required.';
    $objFromNamePhrase = Phrase::getPhrase('research_email_fromname', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strFromName = $objFromNamePhrase ? $objFromNamePhrase->getPhraseValue(ENCODE_HTML) : 'Your Name';
    $objFromAddressPhrase = Phrase::getPhrase('research_email_fromaddress', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strFromAddress = $objFromAddressPhrase ? $objFromAddressPhrase->getPhraseValue(ENCODE_HTML) : 'Your Email Address';
    $objFromPhonePhrase = Phrase::getPhrase('research_email_fromphone', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strFromPhone = $objFromPhonePhrase ? $objFromPhonePhrase->getPhraseValue(ENCODE_HTML) : 'Your Phone Number';
    $objSubjectPhrase = Phrase::getPhrase('research_email_subject', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strSubject = $objSubjectPhrase ? $objSubjectPhrase->getPhraseValue(ENCODE_HTML) : 'Subject';
    $objMessagePhrase = Phrase::getPhrase('research_email_message', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strMessage = $objMessagePhrase ? $objMessagePhrase->getPhraseValue(ENCODE_HTML) : 'Your Message';
    $objSendEmailPhrase = Phrase::getPhrase('research_email_sendemail', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strSendEmail = $objSendEmailPhrase ? $objSendEmailPhrase->getPhraseValue(ENCODE_HTML) : 'Send Email';
    $objCartAppendPhrase = Phrase::getPhrase('research_email_cartappend', PACKAGE_COLLECTIONS, 0, PHRASETYPE_PUBLIC);
    $strCartAppend = $objCartAppendPhrase ? $objCartAppendPhrase->getPhraseValue(ENCODE_HTML) : "Your 'cart' currently holds the following materials.  This list will be appended to your email message.";

    $_ARCHON->PublicInterface->Title = $strEmailTitle;
    $_ARCHON->PublicInterface->addNavigation($_ARCHON->PublicInterface->Title);

    $in_referer = $_REQUEST['referer'] ? $_REQUEST['referer'] : urlencode($_REQUEST['HTTP_REFERER']);

    $repositoryid = $_REQUEST['repositoryid'] ? $_REQUEST['repositoryid'] : 0;



	if(!$_ARCHON->PublicInterface->Templates['core']['Email'])
    {
        $_ARCHON->declareError("Could not display Email: Email template not defined for template set {$_ARCHON->PublicInterface->TemplateSet}.");
    }


    include("header.inc.php");



    if($_ARCHON->Security->userHasAdministrativeAccess())
    {
        include("footer.inc.php");
        return;
    }

    $strName = $_ARCHON->Security->Session->User ? $_ARCHON->Security->Session->User->toString() : '';
    $strName = $_REQUEST['fromname'] ? encode($_REQUEST['fromname'], ENCODE_HTML) : $strName;
    //$strName = encode($strName, ENCODE_HTML);

    $strUserFrom = isset($_ARCHON->Security->Session->User) ? $_ARCHON->Security->Session->User->getString('Email') : '';
    $strFrom = $_REQUEST['fromaddress'] ? encode($_REQUEST['fromaddress'], ENCODE_HTML) : $strUserFrom;
    //$strFrom = encode($strFrom, ENCODE_HTML);

    $strPhone = isset($_ARCHON->Security->Session->User) ? $_ARCHON->Security->Session->User->getString('Phone') : '';
    if($_ARCHON->Security->Session->User->PhoneExtension)
    {
        $strPhone .= " $strExt: " . $_ARCHON->Security->Session->User->getString('PhoneExtension');
    }
    $strPhone = $_REQUEST['fromphone'] ? encode($_REQUEST['fromphone'], ENCODE_HTML) : $strPhone;
    //$strPhone = encode($strPhone, ENCODE_HTML);

	$strPageTitle = strip_tags($_ARCHON->PublicInterface->Title);

	$query_string = htmlspecialchars($_SERVER['QUERY_STRING'], ENT_COMPAT, "UTF-8");

	print "<form action=\"index.php\" accept-charset=\"UTF-8\" method=\"post\">\n";

	$form = "<input type=\"hidden\" name=\"f\" value=\"sendemail\" />\n";
	$form .= "<input type=\"hidden\" name=\"p\" value=\"core/contact\" />\n";
	$form .= "<input type=\"hidden\" name=\"referer\" value=\"$in_referer\" />\n";
	$form .= "<input type=\"hidden\" name=\"query_string\" value=\"$query_string\" />\n";
	$form .= "<input type=\"hidden\" name=\"RepositoryID\" value=\"$repositoryid\" />\n";

	$strRequiredMarker = "<span style=\"color:red\">*</span>";

	$inputs = array();

	$inputs[] = array(
		'strInputLabel' => "<label for=\"name\">$strFromName:</label>",
		'strInputElement' => "<input type=\"text\" name=\"FromName\" id=\"name\" size=\"30\" value=\"$strName\" />",
		'strRequired' => '',
		'template' => 'FieldGeneral',
	);

	$inputs[] = array(
		'strInputLabel' => "<label for=\"email\">$strFromAddress:</label>",
		'strInputElement' => "<input type=\"text\" name=\"FromAddress\" id=\"email\" size=\"25\" value=\"$strFrom\" />",
		'strRequired' => $strRequiredMarker,
		'template' => 'FieldGeneral',
	);

	$inputs[] = array(
		'strInputLabel' => "<label for=\"phone\">$strFromPhone:</label>\n",
		'strInputElement' => "<input type=\"text\" name=\"FromPhone\" id=\"phone\" size=\"20\" value=\"$strPhone\" />",
		'strRequired' => '',
		'template' => 'FieldGeneral',
	);

	$strEncodedSubject = encode($_REQUEST['subject'], ENCODE_HTML);
	$inputs[] = array(
		'strInputLabel' => "<label for=\"subject\">$strSubject:</label>",
		'strInputElement' => "<input type=\"text\" name=\"subject\" id=\"subject\" size=\"40\" value=\"$strEncodedSubject\" />",
		'strRequired' => '',
		'template' => 'FieldGeneral',
	);

	$strEncodedMessage = encode($_REQUEST['message'], ENCODE_HTML);
	$inputs[] = array(
		'strInputLabel' => "<label for=\"message\">$strMessage:</label>",
		'strInputElement' => "<textarea name=\"message\" id=\"message\" cols=\"38\" rows=\"5\">$strEncodedMessage</textarea>",
		'strRequired' => $strRequiredMarker,
		'template' => 'FieldTextArea',
	);

	foreach($inputs as $input)
	{
		$template = array_pop($input);
		$form .= $_ARCHON->PublicInterface->executeTemplate('core', $template, $input);
	}

    if(!$_ARCHON->Error)
    {
        eval($_ARCHON->PublicInterface->Templates['core']['Email']);
    }

	print "</form>\n";
    include('footer.inc.php');
}





function contact_exec()
{
    global $_ARCHON;


    if($_REQUEST['f'] == 'sendemail')
    {


        $_ARCHON->sendEmail($_REQUEST['fromaddress'], $_REQUEST['message'], $_REQUEST['referer'], $_REQUEST['fromname'], $_REQUEST['subject'], $_REQUEST['fromphone'], $_REQUEST['details'], $_REQUEST['detailsfunction'], $_REQUEST['detailsparams'], $_REQUEST['repositoryid']);

        if(!$_ARCHON->Error)
        {
            $msg = "Thank you! Your e-mail has been sent.";
            $uri = strstr($_REQUEST['referer'], '?');
            $location = $uri ? $uri : "index.php";
        }
        else
        {
            //$_REQUEST['f'] = 'email';
            $location = "index.php?" . $_REQUEST['query_string'];

            $params = array_intersect_key($_REQUEST, array_flip(array('fromaddress', 'message', 'fromname', 'subject', 'fromphone')));

        }
    }



    if($_ARCHON->Error)
    {
       $msg = $_ARCHON->clearError();
    }

    if($location)
    {
        $params = $params ? $params : array();
        $_ARCHON->sendMessageAndRedirect($msg, $location, $params);
    }
    else
    {
        $_ARCHON->PublicInterface->Header->Message = $msg;
        contact_initialize();
    }
}


?>
