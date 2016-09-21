<?php
/**
 * Login form
 *
 * @package Archon
 * @author Chris Rishel
 */

isset($_ARCHON) or die();

if ($_ARCHON->config->ForceHTTPS && !$_ARCHON->Security->Session->isSecureConnection())
{
   die('<html><body onLoad="location.href=\'https://' . $_SERVER['HTTP_HOST'] . $_SERVER['REQUEST_URI'] . '\';"></body></html>');
}

$go = $_REQUEST['go'] ? $_REQUEST['go'] : '';
$go = str_replace('f=logout', '', $go);

if($_ARCHON->Security->isAuthenticated())
{
    header("Location: ?p=$go");
}

$PublicPhrasePhraseInputTypeID = $_ARCHON->getPhraseTypeIDFromString('Public Phrase');

$objLoginTitlePhrase = Phrase::getPhrase('login_title', PACKAGE_CORE, 0, $PublicPhrasePhraseInputTypeID);
$strLoginTitle = $objLoginTitlePhrase ? $objLoginTitlePhrase->getPhraseValue(ENCODE_HTML) : 'Login or Register an Account';

$_ARCHON->PublicInterface->Title = $strLoginTitle;
$_ARCHON->PublicInterface->addNavigation($_ARCHON->PublicInterface->Title);

require_once("header.inc.php");

$objSelectOnePhrase = Phrase::getPhrase('register_selectone', PACKAGE_CORE, 0, $PublicPhrasePhraseInputTypeID);
$strSelectOne = $objSelectOnePhrase ? $objSelectOnePhrase->getPhraseValue(ENCODE_HTML) : '(Select One)';

$objLoginPhrase = Phrase::getPhrase('login_login', PACKAGE_CORE, 0, $PublicPhrasePhraseInputTypeID);
$strLogin = $objLoginPhrase ? $objLoginPhrase->getPhraseValue(ENCODE_HTML) : 'Login';
$objPasswordPhrase = Phrase::getPhrase('login_password', PACKAGE_CORE, 0, $PublicPhrasePhraseInputTypeID);
$strPassword = $objPasswordPhrase ? $objPasswordPhrase->getPhraseValue(ENCODE_HTML) : 'Password';
$objRememberMePhrase = Phrase::getPhrase('login_rememberme', PACKAGE_CORE, 0, $PublicPhrasePhraseInputTypeID);
$strRememberMe = $objRememberMePhrase ? $objRememberMePhrase->getPhraseValue(ENCODE_HTML) : 'Remember Me';

$strPageTitle = strip_tags($strLoginTitle);

$strSubmitButton = "<input type=\"submit\" value=\"$strLogin\" class=\"button\" />";

$vars = array();

// Why is the value for this button not internationalized?
$registerButton = "<input type=\"button\" value=\"Register an Account\" onclick=\"location.href='?p=core/register&amp;go=$go';\" />\n";



$inputs[] = array(
	'strInputLabel' => "<label for=\"ArchonLoginFieldA\">$strLogin:</label>",
	'strInputElement' => "<input type=\"text\" id=\"ArchonLoginFieldA\" name=\"ArchonLogin\" value=\"$_REQUEST[login]\" maxlength=\"50\" />",
	'strRequired' => '',
	'template' => 'FieldGeneral',
);

$inputs[] = array(
	'strInputLabel' => "<label for=\"ArchonPasswordFieldA\">$strPassword:</label>",
	'strInputElement' => "<input type=\"password\" id=\"ArchonPasswordFieldA\" name=\"ArchonPassword\" />",
	'strRequired' => '',
	'template' => 'FieldGeneral',
);

$inputs[] = array(
	'strInputLabel' => "<label for=\"RememberMeFieldA\">$strRememberMe:</label>",
	'strInputElement' => "<input type=\"checkbox\" name=\"RememberMe\" id=\"RememberMeFieldA\" value=\"1\" />",
	'strRequired' => '',
	'template' => 'FieldGeneral',
);

$form = "<input type=\"hidden\" name=\"p\" value=\"$_REQUEST[p]\" />\n";

foreach($inputs as $input)
{
	$template = array_pop($input);
	$form .= $_ARCHON->PublicInterface->executeTemplate('core', $template, $input);
}


echo("<form action=\"index.php\" accept-charset=\"UTF-8\" method=\"post\">\n");

if(!$_ARCHON->Error)
{
	eval($_ARCHON->PublicInterface->Templates['core']['Login']);
}

print "</form>\n";
require_once("footer.inc.php");
?>