<?php
/**
 * Profile editing template
 *
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
<h1 id="titleheader"><?php echo($strPageTitle); ?></h1>

<div class='userformbox bground'>

  <p class="center">
    <span class="bold"><?php echo($strPageTitle); ?></span><br />
    <?php echo($strRequired); ?>
  </p>

<?php echo($form); ?>

  <div id="userformsubmit">
    <?php echo($strSubmitButton); ?>
  </div>
<br />
<p class="center"><a href="?p=core/privacy"><?php echo($strPrivacyNote); ?></a></p>
</div>
