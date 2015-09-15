<?php
/**
 * Created by PhpStorm.
 * User: assadmahmood
 * Date: 30/06/15
 * Time: 10:45 AM
 */

namespace Aspose\Words\LoadingAndSaving;


class LoadTxt {

    public static function run($dataDir=null)
    {
        if(is_null($dataDir)) die("Data Directory Undefined");

        $comHelper = new \COM("Aspose.Words.ComHelper");



        // The encoding of the text file is automatically detected.
        $doc = $comHelper->Open($dataDir."/LoadTxt.txt");

        // Save as any Aspose.Words supported format, such as DOCX.
        $doc->Save($dataDir . "/LoadTxt Out.docx");

        print "Text document loaded successfully.\nFile saved at " . $dataDir . "LoadTxt Out.docx" . PHP_EOL;

    }

} 