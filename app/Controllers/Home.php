<?php

namespace App\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory;

class Home extends BaseController
{
    public function index() {
        return view('home');
    }

    public function excelToWord() {
        $file = $this->request->getFile( "excel_file" );

        if ( empty( $file ) ) die( "no file" );

        $file_name = $file->getName();
        $file_extension = array_reverse( explode( ".", $file_name ) )[ 0 ];

        if ( $file_extension !== "xlsx" ) die( "wrong format" );

        $reader = IOFactory::createReader( $file->getLinkTarget() );

        var_dump( $reader );
        die( "done" );

        $spreadsheet = $reader->load( $inputFileName );

        $schdeules = $spreadsheet->getActiveSheet()->toArray();
    }
}
