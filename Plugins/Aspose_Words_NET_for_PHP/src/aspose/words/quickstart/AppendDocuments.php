<?php
/**
 * Created by PhpStorm.
 * User: assadmahmood
 * Date: 30/06/15
 * Time: 10:45 AM
 */

namespace Aspose\Words\QuickStart;


class AppendDocuments {

    public static function run($dataDir=null)
    {
        if(is_null($dataDir)) die("Data Directory Undefined");

        $comHelper = new \COM("Aspose.Words.ComHelper");

        // Load the destination and source documents from disk.
        $dstDoc = $comHelper->Open($dataDir."/TestFile.Destination.doc");
        $srcDoc = $comHelper->Open($dataDir."/TestFile.Source.doc");

        // Append the source document to the destination document while keeping the original formatting of the source document.
        $dstDoc->AppendDocument($srcDoc,1);

        $dstDoc->Save($dataDir . "/TestFile Out.docx");

        echo "Document appended successfully.\nFile saved at " . $dataDir . "TestFile Out.docx" . PHP_EOL;

    }

} 