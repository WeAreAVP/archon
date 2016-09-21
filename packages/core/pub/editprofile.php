<?php
/**
 * Register a user account form
 *
 * @package Archon
 * @author Chris Rishel
 */

isset($_ARCHON) or die();

if(!$_ARCHON->Security->isAuthenticated())
{
    header('Location: index.php?p=');
}

editprofile_initialize();

function editprofile_initialize()
{
	if(!$_REQUEST['f'])
	{
	    editprofile_form();
	}
	else
	{
	    editprofile_exec();
	}
}


function editprofile_form()
{
    global $_ARCHON;

    

    $objMyAccountPhrase = Phrase::getPhrase('myaccount_title', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strMyAccountTitle = $objMyAccountPhrase ? $objMyAccountPhrase->getPhraseValue(ENCODE_HTML) : 'My Account';
    $objEditProfileTitlePhrase = Phrase::getPhrase('editprofile_title', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strEditProfileTitle = $objEditProfileTitlePhrase ? $objEditProfileTitlePhrase->getPhraseValue(ENCODE_HTML) : 'Edit My Profile';

    $_ARCHON->PublicInterface->Title = $strEditProfileTitle;
    $_ARCHON->PublicInterface->addNavigation($strMyAccountTitle, "?p=core/account");
    $_ARCHON->PublicInterface->addNavigation($_ARCHON->PublicInterface->Title, "?p={$_REQUEST['p']}");

    require_once("header.inc.php");
    
    $arrLanguages = $_ARCHON->getAllLanguages();
    $arrCountries = $_ARCHON->getAllCountries();
    
    $UserCountryID = $_REQUEST['countryid'] ? $_REQUEST['countryid'] : $_ARCHON->Security->Session->User->CountryID;
    
    $objSelectOnePhrase = Phrase::getPhrase('selectone', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strSelectOne = $objSelectOnePhrase ? $objSelectOnePhrase->getPhraseValue(ENCODE_HTML) : '(Select One)';
    $objRequiredPhrase = Phrase::getPhrase('requirednotice', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strRequired = $objRequiredPhrase ? $objRequiredPhrase->getPhraseValue(ENCODE_NONE) : 'Fields marked with an asterisk (<span style="color:red">*</span>) are required.';
    $objYesPhrase = Phrase::getPhrase('yes', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strYes = $objYesPhrase ? $objYesPhrase->getPhraseValue(ENCODE_NONE) : 'Yes';
    $objNoPhrase = Phrase::getPhrase('no', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strNo = $objNoPhrase ? $objNoPhrase->getPhraseValue(ENCODE_NONE) : 'No';
    $objSubmitPhrase = Phrase::getPhrase('submit', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strSubmit = $objSubmitPhrase ? $objSubmitPhrase->getPhraseValue(ENCODE_HTML) : 'Submit';
    
    $objEmailPhrase = Phrase::getPhrase('editprofile_email', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strEmail = $objEmailPhrase ? $objEmailPhrase->getPhraseValue(ENCODE_HTML) : 'E-mail';
    $objFirstNamePhrase = Phrase::getPhrase('editprofile_firstname', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strFirstName = $objFirstNamePhrase ? $objFirstNamePhrase->getPhraseValue(ENCODE_HTML) : 'First Name';
    $objLastNamePhrase = Phrase::getPhrase('editprofile_lastname', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strLastName = $objLastNamePhrase ? $objLastNamePhrase->getPhraseValue(ENCODE_HTML) : 'Last Name';
    $objDisplayNamePhrase = Phrase::getPhrase('editprofile_displayname', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strDisplayName = $objDisplayNamePhrase ? $objDisplayNamePhrase->getPhraseValue(ENCODE_HTML) : 'Display Name';
    $objLanguagePhrase = Phrase::getPhrase('editprofile_language', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strLanguage = $objLanguagePhrase ? $objLanguagePhrase->getPhraseValue(ENCODE_HTML) : 'Language';
    $objCountryPhrase = Phrase::getPhrase('editprofile_country', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strCountry = $objCountryPhrase ? $objCountryPhrase->getPhraseValue(ENCODE_HTML) : 'Country';
    $objPasswordPhrase = Phrase::getPhrase('editprofile_password', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strPassword = $objPasswordPhrase ? $objPasswordPhrase->getPhraseValue(ENCODE_HTML) : 'Password';
    $objConfirmPasswordPhrase = Phrase::getPhrase('editprofile_confirmpassword', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strConfirmPassword = $objConfirmPasswordPhrase ? $objConfirmPasswordPhrase->getPhraseValue(ENCODE_HTML) : 'Confirm Password';
    $objPrivacyNotePhrase = Phrase::getPhrase('editprofile_privacynote', PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
    $strPrivacyNote = $objPrivacyNotePhrase ? $objPrivacyNotePhrase->getPhraseValue(ENCODE_HTML) : 'Privacy Note';

	$strPageTitle = strip_tags($_ARCHON->PublicInterface->Title);

	$strUserID = $_ARCHON->Security->Session->User->ID;

	$strRequiredMarker = "<span style=\"color:red\">*</span>";
	$strSubmitButton = "<input type=\"submit\" value=\"$strSubmit\" class=\"button\" />";

	$inputs = array();

	$strEmailValue = $_ARCHON->Security->Session->User->Email;
	$inputs[] = array(
		'strInputLabel' => "<label for=\"EmailField\">$strEmail:</label>",
		'strInputElement' => "<input type=\"text\" id=\"EmailField\" name=\"Email\" value=\"$strEmailValue\" maxlength=\"50\" />",
		'strRequired' => $strRequiredMarker,
		'template' => 'FieldGeneral',
	);


	$strFirstNameValue = $_ARCHON->Security->Session->User->FirstName;
	$inputs[] = array(
		'strInputLabel' => "<label for=\"FirstNameField\">$strFirstName:</label>",
		'strInputElement' => "<input type=\"text\" id=\"FirstNameField\" name=\"FirstName\" value=\"$strFirstNameValue\" maxlength=\"50\" />",
		'strRequired' => $strRequiredMarker,
		'template' => 'FieldGeneral',
	);


	$strLastNameValue = $_ARCHON->Security->Session->User->LastName;
	$inputs[] = array(
		'strInputLabel' => "<label for=\"LastNameField\">$strLastName:</label>",
		'strInputElement' => "<input type=\"text\" id=\"LastNameField\" name=\"LastName\" value=\"$strLastNameValue\" maxlength=\"50\" />",
		'strRequired' => $strRequiredMarker,
		'template' => 'FieldGeneral',
	);


	$strDisplayNameValue = $_ARCHON->Security->Session->User->DisplayName;
	$inputs[] = array(
		'strInputLabel' => "<label for=\"DisplayNameField\">$strDisplayName:</label>",
		'strInputElement' => "<input type=\"text\" id=\"DisplayNameField\" name=\"DisplayName\" value=\"$strDisplayNameValue\" maxlength=\"100\" />",
		'strRequired' => $strRequiredMarker,
		'template' => 'FieldGeneral',
	);


	$inputs[] = array(
		'strInputLabel' => "<label for=\"PasswordField\">$strPassword:</label>",
		'strInputElement' => "<input type=\"password\" id=\"PasswordField\" name=\"Password\" />",
		'strRequired' => '',
		'template' => 'FieldGeneral',
	);



	$inputs[] = array(
		'strInputLabel' => "<label for=\"ConfirmPasswordField\">$strConfirmPassword:</label>",
		'strInputElement' => "<input type=\"password\" id=\"ConfirmPasswordField\" name=\"ConfirmPassword\" />",
		'strRequired' => '',
		'template' => 'FieldGeneral',
	);


	$strLangSelect = <<<EOT
<select id="LanguangeIDField" name="LanguangeID">
		<option value="0">$strSelectOne</option>
EOT;

	if(!empty($arrLanguages))
	{
		foreach($arrLanguages as $objLanguage)
		{
            $selected = ($_ARCHON->Security->Session->User->LanguageID == $objLanguage->ID) ? ' selected="selected"' : '';
			$strLangSelect .= "		<option value=\"$objLanguage->ID\"$selected>" . $objLanguage->toString() . "</option>\n";
		}
	}

	$strLangSelect .= '</select>';

	$inputs[] = array(
		'strInputLabel' => "<label for=\"LanguangeIDField\">$strLanguage:</label>",
		'strInputElement' => $strLangSelect,
		'strRequired' => '',
		'template' => 'FieldGeneral',
	);


	$strCountrySelect = <<<EOT
<select id="CountryIDField" name="CountryID">
		<option value="0">$strSelectOne</option>
EOT;

	if(!empty($arrCountries))
	{
		foreach($arrCountries as $objCountry)
		{
            $selected = ($UserCountryID == $objCountry->ID) ? ' selected="selected"' : '';
			$strCountrySelect .= "		<option value=\"$objCountry->ID\"$selected>" . $objCountry->toString() . "</option>\n";
		}
	}

	$strCountrySelect .= '</select>';

	$inputs[] = array(
		'strInputLabel' => "<label for=\"CountryIDField\">$strCountry:</label>",
		'strInputElement' => $strCountrySelect,
		'strRequired' => $strRequiredMarker,
		'template' => 'FieldGeneral',
	);

    $prevUserProfileFieldCategoryID = 0;
    
    $_ARCHON->Security->Session->User->dbLoadUserProfileFields();

	// Loop through the fields and add them to the form.
    foreach($_ARCHON->Security->Session->User->UserProfileFields as $Key => $objUserProfileField)
    {
    	if(is_natural($Key) && $objUserProfileField->UserEditable && (empty($objUserProfileField->Countries) || isset($objUserProfileField->Countries[$UserCountryID])))
    	{


			// Handle section headings
    	    if($prevUserProfileFieldCategoryID != $objUserProfileField->UserProfileFieldCategoryID)
            {
				$inputs[] = array(
					'strSectionHeading' => $objUserProfileField->UserProfileFieldCategory->toString(),
					'template' => 'FieldCategory',
				);

                $prevUserProfileFieldCategoryID = $objUserProfileField->UserProfileFieldCategoryID;
            }


		$objUserProfileFieldPhrase = Phrase::getPhrase('editprofile_' . strtolower($objUserProfileField->UserProfileField), PACKAGE_CORE, 0, PHRASETYPE_PUBLIC);
	        $strUserProfileField = $objUserProfileFieldPhrase ? $objUserProfileFieldPhrase->getPhraseValue(ENCODE_HTML) : $objUserProfileField->UserProfileField;
	        
	        $required = $objUserProfileField->Required || (isset($objUserProfileField->Countries[$UserCountryID]) && $objUserProfileField->Countries[$UserCountryID]->Required) ? $strRequiredMarker : '';
	        $value = isset($_REQUEST['userprofilefields'][$objUserProfileField->ID]['value']) ? $_REQUEST['userprofilefields'][$objUserProfileField->ID]['value'] : $_ARCHON->Security->Session->User->UserProfileFields[$objUserProfileField->ID]->Value;


	
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

if($id == 'StateProvinceIDField'){
	ob_start();
	print "<pre>";
	var_dump($arrSelectChoices);
	print "</pre>";
	$details = ob_get_contents();
	ob_end_clean();
	//print $details;
}

				$strInput = <<<EOT
      <select id="$id" name="UserProfileFields[$fieldName][Value]">
        <option value="0">$strSelectOne</option>
EOT;

                if(!empty($arrSelectChoices))
                {
                    foreach($arrSelectChoices as $obj)
                    {
                    	if(!property_exists($obj, 'CountryID') || !isset($obj->CountryID) || $obj->CountryID == $UserCountryID)
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

	$strSubmitButton = "<input type=\"submit\" value=\"$strSubmit\" class=\"button\" />";


	echo("<form action=\"index.php\" accept-charset=\"UTF-8\" method=\"post\">\n");


	$form = "<input type=\"hidden\" name=\"p\" value=\"$_REQUEST[p]\" />\n";
	$form .= "<input type=\"hidden\" name=\"f\" value=\"store\" />\n";
	$form .= "<input type=\"hidden\" name=\"id\" value=\"$strUserID\" />\n";

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
		eval($_ARCHON->PublicInterface->Templates['core']['EditProfile']);
	}

	echo("</form>\n");
    require_once("footer.inc.php");
}




function editprofile_exec()
{
    global $_ARCHON;

    $objUser = New User($_REQUEST);
    
    $objTmpUser = New User($_ARCHON->Security->Session->User->ID);
    
    if(!$objTmpUser->dbLoad())
    {
    	$_ARCHON->declareError("Could not store User: There was already an error.");
    	$_REQUEST['f'] = '';
    }
    else
    {
    	$objUser->Login = $objTmpUser->Login;
    	$objUser->RegisterTime = $objTmpUser->RegisterTime;
    	$objUser->Pending = $objTmpUser->Pending;
    	$objUser->PendingHash = $objTmpUser->PendingHash;
//    	$objUser->RepositoryID = $objTmpUser->RepositoryID;
    	$objUser->RepositoryLimit = $objTmpUser->RepositoryLimit;
    	$objUser->Locked = $objTmpUser->Locked;

        //Make sure user stays as a public user
        $objUser->IsAdminUser = 0;

	    if($_REQUEST['f'] == 'store')
	    {
	    	$arrUserProfileFields = $_ARCHON->getAllUserProfileFields();
	    	
	        if(!empty($arrUserProfileFields))
	        {
	        	foreach($arrUserProfileFields as $objUserProfileField)
	        	{
	        		if(!$_ARCHON->Security->userHasAdministrativeAccess() && !$_REQUEST['userprofilefields'][$objUserProfileField->ID]['value'] && ($objUserProfileField->Required || (isset($objUserProfileField->Countries[$objUser->CountryID]) && $objUserProfileField->Countries[$objUser->CountryID]->Required)))
	        		{
	        			$_ARCHON->declareError("Could not store User: Required field $objUserProfileField->UserProfileField is empty.");
	        			$_REQUEST['f'] = '';
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
	            $_REQUEST['f'] = '';
	        }
	        elseif(!$_ARCHON->Error && $objUser->dbStore())
	        {
	            foreach($_REQUEST['userprofilefields'] as $UserProfileFieldID => $arr)
	            {
	            	if(isset($arrUserProfileFields[$UserProfileFieldID]) && $arrUserProfileFields[$UserProfileFieldID]->UserEditable)
	            	{
                        $objUser->dbSetUserProfileField($UserProfileFieldID, $arr['value']);
	            	}
	            }
	
	            $msg = "Profile updated successfully.";
	
	            if(!$_ARCHON->Error)
	            {
                    $location = '?p=core/account';
	            }
	            else
	            {
	            	$location = '?p=core/editprofile';
	            }
	        }
	        else
	        {
	            $_REQUEST['f'] = '';
	        }
	    }
	    else
	    {
	        $_ARCHON->declareError("Unknown Command: {$_REQUEST['f']}");
	        $_REQUEST['f'] = '';
	    }
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
        editprofile_initialize();
    }
}
