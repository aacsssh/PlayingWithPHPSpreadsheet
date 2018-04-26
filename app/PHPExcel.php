<?php

namespace App;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;

class PHPExcel
{
    private $headerCells = [
        'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I'
    ];
    private $headerColumnIndex = '1';
    private $contentColumnInitialIndex = 2;
    private $headerColumnValues = [];

    /**
     * @var PhpOffice\PhpSpreadsheet\Writer\Xls
     */
    private $excelWriter;
    private $activeSheet;

    /**
     * Get XML data into new Excel file
     * 
     * @param  string $filePath  Path of the file
     * @param  string $filename  Name of the download file
     * @return void
     */
    public function xmlToExcel($filePath, $filename)
    {
        $books = simplexml_load_file($filePath);
        $spreadsheet = new Spreadsheet;
        $this->excelWriter = new Xls($spreadsheet);
        $spreadsheet->setActiveSheetIndex(0);
        $this->activeSheet = $spreadsheet->getActiveSheet();

        foreach ($books as $book) {
            $props = get_object_vars($book);
            unset($props['@attributes']);
            $this->setContent($props);
        }

        $this->download('application/vnd.ms-excel', $filename);
    }

    /**
     * Set the content of the excel file
     * @param array $props
     */
    public function setContent($props)
    {
        $titleCount = 0;

        foreach ($props as $title => $value) {
            $cell = $this->headerCells[$titleCount];
            $this->setHeaderColumn($cell, $title);
            $this->activeSheet->setCellValue($cell . "{$this->contentColumnInitialIndex}", $value);
            ++$titleCount;
        }

        ++$this->contentColumnInitialIndex;
    }

    /**
     * Set the header coloumn of the excel file
     * @param string $cell  Cell index of the excel file like A, B, C
     * @param string $value Value of the header column
     */
    public function setHeaderColumn($cell, $value)
    {
        if(in_array($value, $this->headerColumnValues)) {
            return;
        }

        $this->activeSheet->setCellValue($cell . "{$this->headerColumnIndex}", $value);
        $this->headerColumnValues[] = $value;
    }

    /**
     * Download the file
     * @param string $contentType Content type of the file to be downloaded
     * @param string $filename    Name of the file
     */
    public function download($contentType, $filename)
    {
        header("Content-Type: {$contentType}");
        header("Content-Disposition: attachment;filename={$filename}"); /*-- $filename is  xsl filename ---*/
        header('Cache-Control: max-age=0');
        $this->excelWriter->save('php://output');
    }
}
