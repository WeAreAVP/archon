<?php
/**
 * Form sub-template: textareas
 *
 * This file inserts a textarea and its label into a form.
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
	<div class="userformlabel"><?php echo($strInputLabel); echo($strRequired); ?></div>
	<div class="userforminput">
		<?php echo($strInputElement); ?>
	</div>
</div>

