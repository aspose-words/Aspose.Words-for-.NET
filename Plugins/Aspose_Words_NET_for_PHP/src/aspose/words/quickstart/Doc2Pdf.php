<?php
/**
 * Created by PhpStorm.
 * User: assadmahmood
 * Date: 30/06/15
 * Time: 10:45 AM
 */

namespace Aspose\Words\QuickStart;


class Doc2Pdf {

    public static function run($dataDir=null)
    {
        if(is_null($dataDir)) die("Data Directory Undefined");

        $comHelper = new \COM("Aspose.Words.ComHelper");

        $doc = $comHelper->Open($dataDir."/Template.doc");

        // Save the document in PDF format.
        $doc->Save($dataDir . "/Doc2PdfSave Out.pdf");

        print "Document converted to PDF successfully.\nFile saved at " . $dataDir . "Doc2PdfSave Out.pdf" . PHP_EOL;

    }

} 