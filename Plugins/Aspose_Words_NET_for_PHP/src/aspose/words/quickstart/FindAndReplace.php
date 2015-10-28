<?php
/**
 * Created by PhpStorm.
 * User: assadmahmood
 * Date: 30/06/15
 * Time: 10:45 AM
 */

namespace Aspose\Words\QuickStart;


class FindAndReplace {

    public static function run($dataDir=null)
    {
        if(is_null($dataDir)) die("Data Directory Undefined");

        $comHelper = new \COM("Aspose.Words.ComHelper");

        $doc = $comHelper->Open($dataDir."/ReplaceSimple.doc");

        // Check the text of the document
        print "Original document text: " . $doc->Range->Text . PHP_EOL;

        // Replace the text in the document.
        $doc->Range->Replace("_CustomerName_", "James Bond", false, false);

        // Check the replacement was made.
        print "Original document text: " . $doc->Range->Text . PHP_EOL;

        // Save the modified document.
        $doc->Save($dataDir . "/ReplaceSimple Out.doc");

        print "Text found and replaced successfully.\nFile saved at " . $dataDir . "ReplaceSimple Out.doc" . PHP_EOL;

    }

} 