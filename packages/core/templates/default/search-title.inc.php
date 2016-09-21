<?php
/**
 * Search Title template (in case someone needs to change the H1 to something else).
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
