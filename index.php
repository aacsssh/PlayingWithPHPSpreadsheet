<?php

require 'vendor/autoload.php';

error_reporting(E_ERROR | E_PARSE);

use App\PHPExcel;

$excel = new PHPExcel;
$excel->xmlToExcel('books.xml', 'books.xls');