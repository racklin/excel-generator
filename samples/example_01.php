<?php

require '../vendor/autoload.php';

use Racklin\ExcelGenerator\ExcelGenerator;

$pdf = new ExcelGenerator();


$pdf->generate('example_01.json', ["name"=>"rack", "cname"=>"阿土伯", "data"=> [
    ["a"=>"A1", "b"=>"B1"],
    ["a"=>"A2", "b"=>"B2"],
    ["a"=>"A3", "b"=>"B2"],
    ["a"=>"A4", "b"=>"B2"],
    ["a"=>"A5", "b"=>"B2"],
    ["a"=>"A6", "b"=>"B2"],
    ["a"=>"A7", "b"=>"B2"],
    ["a"=>"A8", "b"=>"B2"],
    ["a"=>"A9", "b"=>"B2"],
]], '/tmp/example_01.xlsx', 'F');
