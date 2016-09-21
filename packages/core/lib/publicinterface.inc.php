<?php
abstract class Core_PublicInterface
{
    /**
     * Adds entry to navigation array for use in a "breadcrumbs" style display in the public interface
     *
     * @param string $Title
     * @param string $URL
     * @param boolean $AddToFront
     * @return boolean
     */
    public function addNavigation($Title, $URL = NULL, $AddToFront = false)
    {
        if(!$Title)
        {
            return false;
        }

        $objNavigation->Title = $Title;
        $objNavigation->URL = $URL;

        if($AddToFront)
        {
            array_unshift($this->Navigation, $objNavigation);
        }
        else
        {
            array_push($this->Navigation, $objNavigation);
        }

        return true;
    }




    /**
     * Returns HTML for "breadcrumbs" style navigation display
     *
     * @return string
     */
    public function createNavigation()
    {
        global $_ARCHON;

        if(!empty($this->Navigation))
        {
            $Count = 0;

            foreach($this->Navigation as $objNavigation)
            {
                $Count++;

                $String .= ($objNavigation->URL && count($this->Navigation) > $Count) ? "<a href='$objNavigation->URL'>".$objNavigation->Title."</a>" : $objNavigation->Title;

                if(count($this->Navigation) > $Count)
                {
                    $String .= $_ARCHON->PublicInterface->Delimiter;
                }
            }
        }

        return $String;
    }



    public function outputGoogleAnalyticsCode()
    {
       global $_ARCHON;

       if($_ARCHON->config->GACode)
       {
       ?>
<script type="text/javascript">
  var _gaq = _gaq || [];
  _gaq.push(['_setAccount', '<?php echo($_ARCHON->config->GACode); ?>']);
  _gaq.push(['_trackPageview']);
  (function() {
    var ga = document.createElement('script'); ga.type = 'text/javascript'; ga.async = true;
    ga.src = ('https:' == document.location.protocol ? 'https://ssl' : 'http://www') + '.google-analytics.com/ga.js';
    var s = document.getElementsByTagName('script')[0]; s.parentNode.insertBefore(ga, s);
  })();
</script>
       <?php
       }
    }



    /**
     * Initializes PublicInterface
     *
     * @param string $Theme
     * @param string $TemplateSet
     */
    public function initialize($Theme, $TemplateSet)
    {
        global $_ARCHON;

        if(preg_match('/[\\/\\\\]/u', $Theme) || !file_exists('themes/' . $Theme))
        {
            $Theme = CONFIG_CORE_DEFAULT_THEME;
        }

        $this->Theme = $Theme;

        $this->ImagePath = "themes/$Theme/images";
        if(is_dir("themes/$Theme/js"))
        {
            $this->ThemeJavascriptPath = "themes/$Theme/js";
        }

        if(file_exists('themes/' . $this->Theme . '/init.inc.php'))
        {
            $cwd = getcwd();

            chdir('themes/' . $this->Theme);

            require_once('init.inc.php');

            chdir($cwd);
        }

        $this->TemplateSet = $TemplateSet;
        $this->Templates = $_ARCHON->loadTemplates($this->TemplateSet);
    }

   /**
    * Executes a template.
    *
    * @param string $package Package whose template set contains the template file.
    * @param string $template The name of the template, as registered in the template directory's index.php.
    * @param array $vars An associative array of variable names, which will be extracted and supplied to the template for printout.
    */
	public function executeTemplate($package, $template, $vars)
	{
		global $_ARCHON;
		extract($vars, EXTR_SKIP);

		ob_start();
		eval($this->Templates[$package][$template]);
		$result = ob_get_contents();
		ob_end_clean();

		return $result;
	}



    /**
     * Indicates if toString and getString functions should escape values before returning their string
     *
     * @var boolean
     */
    public $EscapeXML = CONFIG_ESCAPE_XML;

    public $Delimiter = ' -> ';

    public $DisableTheme = false;

    public $ImagePath = NULL;

    public $TemplateSet = NULL;

    public $Title = NULL;

    public $Theme = NULL;

    public $Navigation = array();

    public $Templates = array();

    public $PublicSearchFunctions = array();
}

$_ARCHON->mixClasses('PublicInterface', 'Core_PublicInterface');
?>
