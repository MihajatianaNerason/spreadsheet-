<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;
use App\Http\Controllers\Controller;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class TestController extends Controller
{
    function index() {
        // Creates New Spreadsheet 
        $spreadsheet = new Spreadsheet(); 
        
        // Retrieve the current active worksheet 
        $sheet = $spreadsheet->getActiveSheet(); 
        
        $sheet->setTitle('Sheet 1'); // This is where you set the title 
        $sheet->setCellValue('A1', 'No'); // This is where you set the column header
        $sheet->setCellValue('B1', 'Name');// This is where you set the column header
        $row = 2;// Initialize row counter

        // This is the loop to populate data
        for ($i=1; $i < 5; $i++) { 
            $sheet->setCellValue('A' . $row, $i);
            $sheet->setCellValue('B' . $row, "People ".$i);
            $row++;
        
        } 
        
        // Write an .xlsx file  
        $writer = new Xlsx($spreadsheet); 


        $fileName = "andrana.xlsx";
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Disposition: attachment;filename=\"$fileName\"");
        return $writer->save("php://output");
        
        // Save .xlsx file to the current directory 
        // return $writer->save('gfg.xlsx'); 

    }
}
