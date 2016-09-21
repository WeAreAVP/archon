<?php
/**
 * Footer file for Archon Nouveau theme
 *
 * @package Archon
 * @author Will Martin
 */

isset($_ARCHON) or die();

if($_ARCHON->Script == 'packages/collections/pub/findingaid.php')
{
   require("fafooter.inc.php");
   return;
}

$auth_form = '';

if(!$_ARCHON->Security->isAuthenticated()){

$action = encode($_SERVER['REQUEST_URI'], ENCODE_HTML);
$auth_form = <<<EOT
<div style="display: none">
<div id="userlogin">
	<h3 style="margin-top: 0">Login</h3>
	<form action="$action" accept-charset="UTF-8" method="post">
		<div class="dyad"><label for="ArchonLoginField">Login: <input id="ArchonLoginField" type="text" name="ArchonLogin" size="20" /></label></div>
		<div class="dyad"><label for="ArchonPasswordField">Password: <input id="ArchonPasswordField" type="password" name="ArchonPassword" size="20" /></label></div>
		<div class="dyad">
		<input type="submit" value="Log in" class="button" /><br />
		<label for="RememberMeField"><input id="RememberMeField" type="checkbox" name="RememberMe" value="1" /> Remember me?</label>
		</div>
	</form>
</div>
</div>
EOT;
}

?>


</section>
<footer>
<?php print $auth_form; ?>
<p><a href='http://www.archon.org/'>Archon</a> 3.21 ©2012 <a href="http://www.uiuc.edu/" rel="external">The University of Illinois at Urbana-Champaign</a></p>
<p>Marble background image courtesy of <a href="http://www.spiralgraphics.biz/packs/marble/index.htm?21">Spiral Graphics</a></p>
</footer>
</div>

<?php
$_ARCHON->PublicInterface->outputGoogleAnalyticsCode();
?>
