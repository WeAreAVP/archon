<?php

/**
 * Accession importer script.
 *
 * This script takes .csv files in a defined format and creates a new accession record for each row in the database.
 * A sample csv/excel file is provided in the archon/incoming folder, to show the necessary format.
 *
 * If a creator is defined in the CSV file, the script checks to see if an authority entry exists for the creator,
 * then links to the authority entry.  If no authority entry exists, it makes a new creator authority,
 * then links it to the record.
 *
 * this script does not currently support the import and linking of controlled subject or genre terms.
 *
 * @package Archon
 * @subpackage AdminUI
 * @author Kyle Fox
 */
isset($_ARCHON) or die();

$currentRepositoryID = $_REQUEST['currentrepositoryid'];

$UtilityCode = 'accession_xlsx';

$_ARCHON->addDatabaseImportUtility(PACKAGE_ACCESSIONS, $UtilityCode, '3.21', array('xlsx'), true);

if ($_REQUEST['f'] == 'import-' . $UtilityCode) {
    if (!$_ARCHON->Security->verifyPermissions(MODULE_DATABASE, FULL_CONTROL)) {
        die("Permission Denied.");
    }

    if ($currentRepositoryID <= 0) {
        die("Repository ID required.");
    }

    @set_time_limit(0);

    ob_implicit_flush();

    $arrFiles = $_ARCHON->getAllIncomingFiles(false);

    if (!empty($arrFiles)) {
        $arrLocations = $_ARCHON->getAllLocations();
        foreach ($arrLocations as $objLocation) {
            $arrLocationsMap[encoding_strtolower($objLocation->Location)] = $objLocation->ID;
        }

        $arrMaterialTypes = $_ARCHON->getAllMaterialTypes();
        foreach ($arrMaterialTypes as $objMaterialType) {
            $arrMaterialTypesMap[encoding_strtolower($objMaterialType->MaterialType)] = $objMaterialType->ID;
        }

        $arrProcessingPriorities = $_ARCHON->getAllProcessingPriorities();
        foreach ($arrProcessingPriorities as $objProcessingPriority) {
            $arrProcessingPrioritiesMap[encoding_strtolower($objProcessingPriority->ProcessingPriority)] = $objProcessingPriority->ID;
        }

        $arrExtentUnits = $_ARCHON->getAllExtentUnits();
        foreach ($arrExtentUnits as $objExtentUnit) {
            $arrExtentUnitsMap[encoding_strtolower($objExtentUnit->ExtentUnit)] = $objExtentUnit->ID;
        }

        $CreatorTypeID = $_ARCHON->getCreatorTypeIDFromString('Personal Name');

        $arrLanguages = $_ARCHON->getAllLanguages();
        foreach ($arrLanguages as $objLanguage) {
            $arrLanguagesMap[encoding_strtolower($objLanguage->LanguageShort)] = $objLanguage->ID;
        }

        foreach ($arrFiles as $Filename => $strCSV) {
            echo("Parsing file $Filename...$strCSV<br><br>\n\n");

            // Remove byte order mark if it exists.
            $strCSV = ltrim($strCSV, "\xEF\xBB\xBF");

            $arrAllData = readExcel($strCSV);

            // ignore first line?
            foreach ($arrAllData as $key => $arrData) {
                if (!empty($key) && !empty($arrData)) {
                    $objAccession = new Accession();
                    $objAccession->Enabled = 0;
                    if (!empty($arrData[0]['date'])) {
                        $date = getFormattedDate($arrData[0]['date']);
                        $objAccession->AccessionDateMonth = $date[0];
                        $objAccession->AccessionDateDay = $date[1];
                        $objAccession->AccessionDateYear = $date[2];
                    }
                    $log_no = str_pad($arrData[0]['log'], 4, '0', STR_PAD_LEFT);
                    $objAccession->Identifier = $log_no;
                    if (!empty($arrData[0]['source'])) {
                        $objAccession->Title = $arrData[0]['source'] . ' | ' . $log_no . " | [" . $arrData[0]['archives'] . ']';
                    } else {
                        $objAccession->Title = '| ' . $log_no . " | [" . $arrData[0]['archives'] . "]";
                    }
                    $objAccession->ReceivedExtent = count($arrData);
                    $ReceivedExtentUnit = 'Boxes (General)';
                    $objAccession->ReceivedExtentUnitID = $arrExtentUnitsMap[encoding_strtolower($ReceivedExtentUnit)] ? $arrExtentUnitsMap[encoding_strtolower($ReceivedExtentUnit)] : 0;
                    if (!$objAccession->ReceivedExtentUnitID && $ReceivedExtentUnit) {
                        echo("Extent Unit $ReceivedExtentUnit not found!<br>\n");
                    }

                    if (!in_array(strtolower($arrData[0]['range']), array("p", "d", "c"))) {
                        $objAccession->UnprocessedExtent = count($arrData);
                        $UnprocessedExtentUnit = 'Boxes (General)';
                        $objAccession->UnprocessedExtentUnitID = $arrExtentUnitsMap[encoding_strtolower($UnprocessedExtentUnit)] ? $arrExtentUnitsMap[encoding_strtolower($UnprocessedExtentUnit)] : 0;
                        if (!$objAccession->UnprocessedExtentUnitID && $UnprocessedExtentUnit) {
                            echo("Extent Unit $UnprocessedExtentUnit not found!<br>\n");
                        }
                    }
                    // donor ???
//                    $objAccession->Donor = $arrData[0]['source'];

                    $description = generateDescription($arrData);
                    $objAccession->PhysicalDescription = $description;

                    $objAccession->dbStore();
                    if (!$objAccession->ID) {
                        echo("Error storing accession $objAccession->Title: {$_ARCHON->clearError()}<br>\n");
                        continue;
                    }
                    if ($objAccession->ID) {
                        foreach ($arrData as $_value) {
                            $LocationContent = $_value['box_no'];
                            $objLocationEntry = NULL;
                            if ($LocationContent) {
                                $objLocationEntry = New AccessionLocationEntry();
                                $Location = "Archives Stacks";
                                $objLocationEntry->LocationID = $arrLocationsMap[encoding_strtolower($Location)] ? $arrLocationsMap[encoding_strtolower($Location)] : 0;
                                if ($objLocationEntry->LocationID != 0) {
                                    $objLocationEntry->AccessionID = $objAccession->ID;
                                    $objLocationEntry->Content = $LocationContent;
                                    $objLocationEntry->RangeValue = $_value['range'];
                                    $objLocationEntry->Shelf = $_value['shelf'];
                                    if (!empty($_value['section']))
                                        $objLocationEntry->Section = $_value['section'];
                                    $objLocationEntry->Extent = 1;
                                    $LocationEntryExtentUnit = 'Boxes (General)';
                                    $objLocationEntry->ExtentUnitID = $arrExtentUnitsMap[encoding_strtolower($LocationEntryExtentUnit)] ? $arrExtentUnitsMap[encoding_strtolower($LocationEntryExtentUnit)] : 0;
                                    if (!$objLocationEntry->ExtentUnitID && $LocationEntryExtentUnit) {
                                        echo("Extent Unit $LocationEntryExtentUnit not found!<br>\n");
                                    }
                                    if (!$objLocationEntry->dbStore()) {
                                        echo("Error relating LocationEntry to accession: {$_ARCHON->clearError()}<br>\n");
                                    }
                                } else {
                                    echo("Location $Location not found!<br>\n");
                                }
                            }
                        }
                    }

                    if ($objAccession->ID) {
                        echo("Imported {$objAccession->Title}.<br><br>\n\n");
                    }

                    flush();
                }
//                }
            }
        }

        echo("All files imported!");
    }
}

function generateDescription($data) {
    $description = '';
    foreach ($data as $value) {
        if (!empty($value['section']) && !empty($value['shelf'])) {
            $description .= 'Box ' . $value['box_no'] . ', ' . $value['description'] . ' (Range ' . $value['range'] . ', Section ' . $value['section'] . ', Shelf ' . $value['shelf'] . ');' . PHP_EOL;
        } else if (!empty($value['section'])) {
            $description .= 'Box ' . $value['box_no'] . ', ' . $value['description'] . ' (Range ' . $value['range'] . ', Section ' . $value['section'] . ');' . PHP_EOL;
        }else if (!empty($value['shelf'])) {
            $description .= 'Box ' . $value['box_no'] . ', ' . $value['description'] . ' (Range ' . $value['range'] . ', Shelf ' . $value['shelf'] . ');' . PHP_EOL;
        } else {
            $description .= 'Box ' . $value['box_no'] . ', ' . $value['description'] . ' (Range ' . $value['range'] . ');' . PHP_EOL;
        }
    }
    return $description;
}

function getFormattedDate($data) {
    $d = $m = $y = '';
    $date = split('/', str_replace('-', '/', $data));
    if (count($date) == 3) {
        $month = \DateTime::createFromFormat('m', $date[0]);
        $m = $month->format('m');
        $day = \DateTime::createFromFormat('d', $date[1]);
        $d = $day->format('d');
        if ($date[2] <= 99) {
            $year = \DateTime::createFromFormat('y', $date[2]);
            $y = $year->format('Y');
        } else {
            $y = $date[2];
        }
    } else if (count($date) > 0) {
        $month = \DateTime::createFromFormat('m', $date[0]);
        $m = $month->format('m');
        if ($date[1] <= 99) {
            $year = \DateTime::createFromFormat('y', $date[1]);
            $y = $year->format('Y');
        } else {
            $y = $date[1];
        }
    }
    return array($m, $d, $y);

    // return date as month day year
}

?>