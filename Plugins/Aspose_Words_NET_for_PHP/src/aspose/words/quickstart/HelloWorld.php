<?php
/**
 * Created by PhpStorm.
 * User: assadmahmood
 * Date: 30/06/15
 * Time: 10:45 AM
 */

namespace Aspose\Words\QuickStart;


class HelloWorld {

    public static function run($dataDir=null)
    {
        if(is_null($dataDir)) die("Data Directory Undefined");

        $doc = new \COM("Aspose.Words.Document");

        $builder = new \COM("Aspose.Words.DocumentBuilder");

        $builder->Document = $doc;

        $builder->Write("Hello world!");

        $doc->Save($dataDir . "/HelloWorld Out.docx");
    }

} 