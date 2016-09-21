<?php
/**
 * Register a user account form
 *
 * @package Archon
 * @author Chris Rishel
 */

isset($_ARCHON) or die();

if($_ARCHON->Security->isAuthenticated())
{
    header('Location: index.php?p=');
}

register_initialize();

function register_initialize()
{
	if(!$_REQUEST['f'])
	{
	    register_form();
	}
	else
	{
	    register_exec();
	}
}





function register_form()
{
    global $_ARCHON;




    $objRegisterTitlePhrase = Phrase::getPhrase('register_registertitle', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strRegisterTitle = $objRegisterTitlePhrase ? $objRegisterTitlePhrase->getPhraseValue(ENCODE_HTML) : 'Register an Account';

    $_ARCHON->PublicInterface->Title = $strRegisterTitle;
    $_ARCHON->PublicInterface->addNavigation($_ARCHON->PublicInterface->Title);

    $arrCountries = $_ARCHON->getAllCountries();

    $objSelectOnePhrase = Phrase::getPhrase('register_selectone', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strSelectOne = $objSelectOnePhrase ? $objSelectOnePhrase->getPhraseValue(ENCODE_HTML) : '(Select One)';
    $objRequiredPhrase = Phrase::getPhrase('editprofile_required', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strRequired = $objRequiredPhrase ? $objRequiredPhrase->getPhraseValue(ENCODE_NONE) : 'Fields marked with an asterisk (<span style="color:red">*</span>) are required.';
    $objYesPhrase = Phrase::getPhrase('yes', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strYes = $objYesPhrase ? $objYesPhrase->getPhraseValue(ENCODE_NONE) : 'Yes';
    $objNoPhrase = Phrase::getPhrase('no', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strNo = $objNoPhrase ? $objNoPhrase->getPhraseValue(ENCODE_NONE) : 'No';

    $objCountryPhrase = Phrase::getPhrase('register_country', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strCountry = $objCountryPhrase ? $objCountryPhrase->getPhraseValue(ENCODE_HTML) : 'Select Your Country';
    $objLoginPhrase = Phrase::getPhrase('register_login', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strLogin = $objLoginPhrase ? $objLoginPhrase->getPhraseValue(ENCODE_HTML) : 'Login';
    $objEmailPhrase = Phrase::getPhrase('register_email', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strEmail = $objEmailPhrase ? $objEmailPhrase->getPhraseValue(ENCODE_HTML) : 'E-mail';
    $objFirstNamePhrase = Phrase::getPhrase('register_firstname', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strFirstName = $objFirstNamePhrase ? $objFirstNamePhrase->getPhraseValue(ENCODE_HTML) : 'First Name';
    $objLastNamePhrase = Phrase::getPhrase('register_lastname', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strLastName = $objLastNamePhrase ? $objLastNamePhrase->getPhraseValue(ENCODE_HTML) : 'Last Name';
    $objPasswordPhrase = Phrase::getPhrase('register_password', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strPassword = $objPasswordPhrase ? $objPasswordPhrase->getPhraseValue(ENCODE_HTML) : 'Password';
    $objConfirmPasswordPhrase = Phrase::getPhrase('register_confirmpassword', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strConfirmPassword = $objConfirmPasswordPhrase ? $objConfirmPasswordPhrase->getPhraseValue(ENCODE_HTML) : 'Confirm Password';
    $objSubmitPhrase = Phrase::getPhrase('register_submit', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strSubmit = $objSubmitPhrase ? $objSubmitPhrase->getPhraseValue(ENCODE_HTML) : 'Submit';
    $objThankYouPhrase = Phrase::getPhrase('register_thankyou', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strThankYou = $objThankYouPhrase ? $objThankYouPhrase->getPhraseValue(ENCODE_HTML) : 'Thank you for registering.';
    $objPrivacyNotePhrase = Phrase::getPhrase('register_privacynote', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strPrivacyNote = $objPrivacyNotePhrase ? $objPrivacyNotePhrase->getPhraseValue(ENCODE_HTML) : 'Privacy Note';

	$strPageTitle = strip_tags($_ARCHON->PublicInterface->Title);

	$strCountrySelect = <<<EOT
<select id="CountryIDField" name="CountryID">
		<option value="0">$strSelectOne</option>
EOT;

	if(!empty($arrCountries))
	{
		foreach($arrCountries as $objCountry)
		{
			$selected = ($_ARCHON->Security->Session->User->CountryID == $objCountry->ID) ? ' selected="selected"' : '';

			$strCountrySelect .= "		<option value=\"$objCountry->ID\"$selected>" . $objCountry->toString() . "</option>\n";
		}
	}

	$strCountrySelect .= '</select>';

	$inputs = array();

$strRequiredMarker = "<span style=\"color:red\">*</span>";
$strSubmitButton = "<input type=\"submit\" value=\"$strSubmit\" class=\"button\" />";



$inputs[] = array(
	'strInputLabel' => "<label for=\"CountryIDField\">$strCountry:</label>",
	'strInputElement' => $strCountrySelect,
	'strRequired' => '',
	'template' => 'FieldGeneral',
);

$inputs[] = array(
	'strInputLabel' => "<label for=\"EmailField\">$strEmail:</label>",
	'strInputElement' => "<input type=\"text\" id=\"EmailField\" name=\"Email\" value=\"$_REQUEST[email]\" maxlength=\"50\" />",
	'strRequired' => $strRequiredMarker,
	'template' => 'FieldGeneral',
);

$inputs[] = array(
	'strInputLabel' => "<label for=\"FirstNameField\">$strFirstName:</label>",
	'strInputElement' => "<input type=\"text\" id=\"FirstNameField\" name=\"FirstName\" value=\"$_REQUEST[firstname]\" maxlength=\"50\" />",
	'strRequired' => $strRequiredMarker,
	'template' => 'FieldGeneral',
);

$inputs[] = array(
	'strInputLabel' => "<label for=\"LastNameField\">$strLastName:</label>",
	'strInputElement' => "<input type=\"text\" id=\"LastNameField\" name=\"LastName\" value=\"$_REQUEST[lastname]\" maxlength=\"50\" />",
	'strRequired' => $strRequiredMarker,
	'template' => 'FieldGeneral',
);

$inputs[] = array(
	'strInputLabel' => "<label for=\"PasswordField\">$strPassword:</label>",
	'strInputElement' => "<input type=\"password\" id=\"PasswordField\" name=\"Password\" />",
	'strRequired' => $strRequiredMarker,
	'template' => 'FieldGeneral',
);

$inputs[] = array(
	'strInputLabel' => "<label for=\"ConfirmPasswordField\">$strConfirmPassword:</label>",
	'strInputElement' => "<input type=\"password\" id=\"ConfirmPasswordField\" name=\"ConfirmPassword\" />",
	'strRequired' => $strRequiredMarker,
	'template' => 'FieldGeneral',
);
  
    $arrUserProfileFields = $_ARCHON->getAllUserProfileFields();

    $prevUserProfileFieldCategoryID = 0;
    
    foreach($arrUserProfileFields as $Key => $objUserProfileField)
    {
        if(is_natural($Key) && $objUserProfileField->UserEditable && (empty($objUserProfileField->Countries) || isset($objUserProfileField->Countries[$_REQUEST['countryid']])))
        {
            if($prevUserProfileFieldCategoryID != $objUserProfileField->UserProfileFieldCategoryID)
            {
            	$objUserProfileField->dbLoad();

				$inputs[] = array(
					'strSectionHeading' => $objUserProfileField->UserProfileFieldCategory->toString(),
					'template' => 'FieldCategory',
				);

                $prevUserProfileFieldCategoryID = $objUserProfileField->UserProfileFieldCategoryID;
            }
            
            $objUserProfileFieldPhrase = Phrase::getPhrase($objUserProfileField->UserProfileField, PACKAGE_CORE, MODULE_PUBLICUSERS, PHRASETYPE_ADMIN);
            $strUserProfileField = $objUserProfileFieldPhrase ? $objUserProfileFieldPhrase->getPhraseValue(ENCODE_HTML) : $objUserProfileField->UserProfileField;
            
            $required = $objUserProfileField->Required || (isset($objUserProfileField->Countries[$_REQUEST['countryid']]) && $objUserProfileField->Countries[$_REQUEST['countryid']]->Required) ? '<span style="color:red">*</span>' : '';
            $value = isset($_REQUEST['userprofilefields'][$objUserProfileField->ID]['value']) ? $_REQUEST['userprofilefields'][$objUserProfileField->ID]['value'] : $objUserProfileField->DefaultValue;






            if($objUserProfileField->InputType == 'radio')
            {
				if($value)
				{
					$checkedYes = ' checked="checked"';
					$checkedNo = '';
				}
				else
				{
					$checkedYes = '';
					$checkedNo = ' checked="checked"';
				}

				$idYes = $objUserProfileField->UserProfileField.'Yes';
				$idNo = $objUserProfileField->UserProfileField.'No';

				$inputName = "UserProfileFields[".$objUserProfileField->ID."][Value]";

				$radioButtons = <<<EOT
				<fieldset>
					<legend>$strUserProfileField $required</legend>
					<label for="$idYes"><input type="radio" id="$idYes" name="$inputName" value="1"$checkedYes />$strYes</label>
					<label for="$idNo"><input type="radio" id="$idNo" name="$inputName" value="0"$checkedNo  />$strNo</label>
				</fieldset>
EOT;

				$inputs[] = array(
					'strInput' => $radioButtons,
					'template' => 'radio',
				);

            }






            elseif($objUserProfileField->InputType == 'select')
            {
                $arrSelectChoices = call_user_func(array($_ARCHON, $objUserProfileField->ListDataSource));
				$id = $objUserProfileField->UserProfileField.'Field';
				$fieldName = $objUserProfileField->ID;

				$strInput = <<<EOT
      <select id="$id" name="UserProfileFields[$fieldName][Value]">
        <option value="0">$strSelectOne</option>
EOT;

                if(!empty($arrSelectChoices))
                {
                    foreach($arrSelectChoices as $obj)
                    {
                    	if(!property_exists($obj, 'CountryID') || !isset($obj->CountryID) || $obj->CountryID == $_REQUEST['countryid'])
                    	{
                            $selected = ($value == $obj->ID) ? ' selected="selected"' : '';
                            $strInput .= "        <option value=\"$obj->ID\"$selected>" . $obj->toString() . "</option>";
                    	}
                    }
                }

			$strInput .= "      </select>";

				$inputs[] = array(
					'strInputLabel' => "<label for=\"$id\">$strUserProfileField:</label>",
					'strInputElement' => $strInput,
					'strRequired' => $required,
					'template' => 'FieldGeneral',
				);
            }






            elseif($objUserProfileField->InputType == 'textarea')
            {
				$id = $objUserProfileField->UserProfileField.'Field';
				$fieldName = $objUserProfileField->ID;
				$size = $objUserProfileField->Size;

				$inputs[] = array(
					'strInputLabel' => "<label for=\"$id\">$strUserProfileField:</label>",
					'strInputElement' => "<textarea id=\"$id\" name=\"UserProfileFields[$fieldName][Value]\" rows=\"$size\" cols=\"50\">$value</textarea>",
					'strRequired' => $required,
					'template' => 'FieldTextArea',
				);
            }






            elseif($objUserProfileField->InputType == 'textfield' || $objUserProfileField->InputType == 'timestamp')
    		{
                if($value && is_natural($value) && $objUserProfileField->InputType == 'timestamp')
                {
                    $value = date(CONFIG_CORE_DATE_FORMAT, $value);
                }

				$id = $objUserProfileField->UserProfileField.'Field';
				$fieldName = $objUserProfileField->ID;
				$size = $objUserProfileField->Size;
				$maxLength = $objUserProfileField->MaxLength;

				$inputs[] = array(
					'strInputLabel' => "<label for=\"$id\">$strUserProfileField:</label>",
					'strInputElement' => "<input type=\"text\" id=\"$id\" name=\"UserProfileFields[$fieldName][Value]\" value=\"$value\" size=\"$size\" maxlength=\"$maxLength\" />",
					'strRequired' => $required,
					'template' => 'FieldGeneral',
				);
    		}
    	}
    }


    if(!$_REQUEST['countryid'])
    {
		// No country ID, so replace the entire form with the country selector.
		$inputs = array_slice($inputs, 0, 1);
    }
	else
	{
		// A country ID is available, so remove the country selector.
		//$inputs = array_slice($inputs, 1);
		array_shift($inputs);
	}

    require_once("header.inc.php");
	echo("<form action=\"index.php\" accept-charset=\"UTF-8\" method=\"post\">\n");

	$form = "<input type=\"hidden\" name=\"p\" value=\"$_REQUEST[p]\" />\n";

	if($_REQUEST['countryid'])
	{
		$form .= "<input type=\"hidden\" name=\"f\" value=\"store\" />\n";
		$form .= "<input type=\"hidden\" name=\"CountryID\" value=\"$_REQUEST[countryid]\" />\n";
	}

	foreach($inputs as $input)
	{
		$template = array_pop($input);
		if($template != 'radio')
		{
			$form .= $_ARCHON->PublicInterface->executeTemplate('core', $template, $input);
		}
		else
		{
			$form .= $input['strInput'];
		}
	}

	if(!$_ARCHON->Error)
	{
		eval($_ARCHON->PublicInterface->Templates['core']['Register']);
	}

	echo("</form>");
    require_once("footer.inc.php");
}

function register_exec()
{
    global $_ARCHON;

    $objUser = New User($_REQUEST);

    if($_REQUEST['f'] == 'store')
    {
    	$objUser->ID = 0;
        $objUser->IsAdminUser = 0;
    	$objUser->DisplayName = "{$objUser->FirstName} {$objUser->LastName}";
    	$objUser->Login = $objUser->Email;

    	$arrUserProfileFields = $_ARCHON->getAllUserProfileFields();
    	$arrPatterns = $_ARCHON->getAllPatterns();
    	
        if(!empty($arrUserProfileFields))
        {
        	foreach($arrUserProfileFields as $objUserProfileField)
        	{
        		if(!$_REQUEST['userprofilefields'][$objUserProfileField->ID]['value'] && ($objUserProfileField->Required || (isset($objUserProfileField->Countries[$_REQUEST['countryid']]) && $objUserProfileField->Countries[$_REQUEST['countryid']]->Required)))
        		{
        			$_ARCHON->declareError("Could not store User: Required field $objUserProfileField->UserProfileField is empty.");
        		}
        		
        		if($_REQUEST['userprofilefields'][$objUserProfileField->ID]['value'] && $objUserProfileField->PatternID)
        		{
        			if(!$arrPatterns[$objUserProfileField->PatternID]->match($_REQUEST['userprofilefields'][$objUserProfileField->ID]['value']))
        			{
        				$_ARCHON->declareError("Could not store User: '{$_REQUEST['userprofilefields'][$objUserProfileField->ID]['value']}' is not a valid $objUserProfileField->UserProfileField.");
        			}
        		}
                
                if($_REQUEST['userprofilefields'][$objUserProfileField->ID]['value'] && $objUserProfileField->InputType == 'timestamp' && !is_natural($_REQUEST['userprofilefields'][$objUserProfileField->ID]['value']))
                {
                    if(($timeValue = strtotime($_REQUEST['userprofilefields'][$objUserProfileField->ID]['value'])) === false)
                    {
                        $_ARCHON->declareError("Could not store User: strtotime() unable to parse value '{$_REQUEST['userprofilefields'][$objUserProfileField->ID]['value']}'.");
                    }
                    else
                    {
                        $_REQUEST['userprofilefields'][$objUserProfileField->ID]['value'] = $timeValue;
                    }
                }
        	}
        }
        
        if($_REQUEST['password'] != $_REQUEST['confirmpassword'])
        {
            $_ARCHON->declareError("Could not store User: Passwords do not match.");
        }
        elseif(!$_ARCHON->Error && $objUser->dbStore())
        {
        	foreach($_REQUEST['userprofilefields'] as $UserProfileFieldID => $arr)
            {
                $objUser->dbSetUserProfileField($UserProfileFieldID, $arr['value']);
            }
            
            if(CONFIG_CORE_VERIFY_PUBLIC_ACCOUNTS)
            {
                $msg = "Thank you for registering!  An E-mail has been sent to {$objUser->Email} with information about how to confirm your account.";
            }
            else
            {
                $msg = "Thank you for registering!";
                $_ARCHON->Security->verifyCredentials($objUser->Login, $_REQUEST['password']);
            }

            $location = '?p=';
        }
        else
        {
            $_REQUEST['f'] = '';
        }
    }
    else if($_REQUEST['f'] == 'activate')
    {
        $objVerifyUser = New User($_REQUEST['id']);
        $UserLoaded = $objVerifyUser->dbLoad();

        if($UserLoaded && !$objVerifyUser->Pending)
        {
        	$msg = "Account already activated.  Please log in.";
        }
        elseif($UserLoaded && $objVerifyUser->Pending && $objVerifyUser->dbActivate($_REQUEST['v']))
        {
            $msg = "Account activated.  You may now log in.";
        }
        else
        {
            $_ARCHON->declareError("Account activation failed.");
            $_REQUEST['f'] = '';
        }

        $location = '?p=';
    }
    else
    {
        $_ARCHON->declareError("Unknown Command: {$_REQUEST['f']}");
        $_REQUEST['f'] = '';
    }

    if($_ARCHON->Error)
    {
       $msg = $_ARCHON->clearError();
    }

    if($location)
    {
        $_ARCHON->sendMessageAndRedirect($msg, $location);
    }
    else
    {
        $_ARCHON->PublicInterface->Header->Message = $msg;
        register_initialize();
    }
}