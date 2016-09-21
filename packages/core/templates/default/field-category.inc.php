<?php
/**
 * Form sub-template: section headings
 *
 * This file inserts a section heading into a form (such as the registration form)
 * when the section head ID changes from one category to the next.
 *
 * The Archon API is available through the variable:
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
?>
	<div class="userformpair">
	<div class="userformlabel"><b><?php echo($strSectionHeading); ?></b></div>
	<div class="userforminput">&nbsp;</div>
	</div>
