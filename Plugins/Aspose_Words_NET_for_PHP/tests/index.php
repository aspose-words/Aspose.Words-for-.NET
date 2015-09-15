<?php
/**
 * Created by PhpStorm.
 * User: assadmahmood
 * Date: 15/09/15
 * Time: 1:47 PM
 */

require_once __DIR__ . '/../vendor/autoload.php'; // Autoload files using Composer autoload
use Aspose\Words\QuickStart\HelloWorld;
use Aspose\Words\QuickStart\FindAndReplace;
use Aspose\Words\QuickStart\Doc2Pdf;
//use Aspose\Words\QuickStart\AppendDocuments;
use Aspose\Words\LoadingAndSaving\LoadTxt;

//print "Running Aspose\\Words\\QuickStart\\HelloWorld::run()" . PHP_EOL;
//HelloWOrld::run(__DIR__ . '/data/QuickStart/HelloWOrld');

//print "Running Aspose\\Words\\QuickStart\\FindAndReplace::run()" . PHP_EOL;
//FindAndReplace::run(__DIR__ . '/data/QuickStart/FindAndReplace');

//print "Running Aspose\\Words\\QuickStart\\Doc2Pdf::run()" . PHP_EOL;
//Doc2Pdf::run(__DIR__ . '/data/QuickStart/Doc2Pdf');

//print "Running Aspose\\Words\\QuickStart\\AppendDocuments::run()" . PHP_EOL;
//AppendDocuments::run(__DIR__ . '/data/QuickStart/AppendDocuments');

print "Running Aspose\\Words\\LoadingAndSaving\\LoadTxt::run()" . PHP_EOL;
LoadTxt::run(__DIR__ . '/data/LoadingAndSaving/LoadTxt');


