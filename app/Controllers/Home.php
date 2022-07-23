<?php

namespace App\Controllers;

use PhpOffice\PhpSpreadsheet\IOFactory as PhpExcelIOFactory;
use PhpOffice\PhpWord\IOFactory as PhpWordIOFactory;
use PhpOffice\PhpWord\PhpWord;

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

        if ( $file_extension !== "xlsx" && $file_extension !== "csv" ) die( "wrong format" );

        // $this->request->getFiles()["excel_file"]->getLinkTarget()

        if ( $file_extension === "xlsx" ) $file_extension = "Xlsx";
        if ( $file_extension === "csv" ) $file_extension = "Csv";

        $reader = PhpExcelIOFactory::createReader( $file_extension );

        $spreadsheet = $reader->load( $file );

        $rows_and_columns = $spreadsheet->getActiveSheet()->toArray( null, true, true, true );

        $phpWord = new PhpWord();

        $phpWord->setDefaultFontSize(14);
        $phpWord->setDefaultFontName('B Nazanin');
        
        $sectionStyle = new \PhpOffice\PhpWord\Style\Section();
        $sectionStyle->setColsNum( 2 );
        $sectionStyle->setColsSpace( 400 );
        $sectionStyle->setMarginLeft( 400 );
        $sectionStyle->setMarginRight( 400 );
        // $sectionStyle->setMarginTop( 400 );
        // $sectionStyle->setMarginBottom( 400 );

        $section = $phpWord->addSection( $sectionStyle );

        $phpWord->addParagraphStyle("pStyler", array( "align" => "right" ) );

        

        // $fontStyle = new \PhpOffice\PhpWord\Style\Font();
        // $fontStyle->setRTL( true );
        // var_dump( $fontStyle->isRTL() );

        foreach( $rows_and_columns as $index => $row ) {

            if ( $index === 1 ) continue;
            $sender = "هابینو";
            $full_address = "آدرس ";
            $post_code = "کد پستی   ";
            $mobile = "تلفن   ";
            $fullname = "گیرنده ";
            $order_status = "وضعیت سفارش ";
            $order_date = "تاریخ سفارش ";
            $post_method = "";
            $order_id = "شماره سفارش ";
            $city = "";
            $state = "";
            $address = "";

            foreach( $row as $idx => $column ) {
                
                // get address
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: city" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: city"
                ) $city = $column;

                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: state" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: state"
                ) $state = $column;

                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: street address (full)" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: street address (full)"
                ) $address = $column;
                // get post code
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: zip code" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: zip code"
                ) {
                    if ( trim( $post_method ) === "تیپاکس" )
                    $post_code .= "0000000000";
                    else
                    $post_code .= $column;
                }

                // get mobile
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: phone number" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: phone number"
                ) $mobile .= $column;

                // get order id
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "order id"
                ) $order_id .= $column;

                // get mobile
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: full name" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: full name"
                ) $fullname .= $column;
                
                // get order status
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "order status"
                ) $order_status .= $column;
                
                // get order date
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "order date"
                ) $order_date .= $column;
                
                // get post method
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping method"
                ) $post_method .= $column;

            }

            // $section->addText( "{$sender}" );
            $full_address .= $state . " - " . $city . " - " . $address;
            $txt1 = $section->addText( "{$fullname}  {$order_id}", [], array('align' => 'right') );
            $txt2 = $section->addText( "{$full_address}", [], array('align' => 'right') );
            $txt3 = $section->addText( "{$post_code}    {$mobile}", [], array('align' => 'right') );

            // $txt3->setFontStyle( $fontStyle, $fontStyle );
        }

        $objWriter = PhpWordIOFactory::createWriter( $phpWord, 'Word2007' );
        $file_word_name = "Addresses.docx";
        $objWriter->save( FCPATH . $file_word_name );

        echo "<a href='" . base_url( $file_word_name ) . "' id='download' >donwload file</a>
            <script>
                document.getElementById( 'download' ).click();
               // setTimeout( () => { window.location.href='" . base_url( '/' ) . "' }, 500 );
            </script>";
        echo "<hr/>";
        
        die( "done" );

    }

    public function writeWord( $total_rows, $total_cells, $cell_text ) {
        $phpWord = new PhpWord();
        $section = $phpWord->addSection();
        
        // $header = $section->addHeader();
        // $header->addText('This is my fabulous header!');
        
        // $footer = $section->addFooter();
        // $footer->addText('Footer text goes here.');
        
        // $textrun = $section->addTextRun();
        // $textrun->addText('Some text. ');
        // $textrun->addText('And more Text in this Paragraph.');
        
        // $textrun = $section->addTextRun();
        // $textrun->addText('New Paragraph! ', ['bold' => true]);
        // $textrun->addText('With text...', ['italic' => true]);
        
        $section->addText( 'Basic table', ['size' => 16, 'bold' => true ] );
        
        $table = $section->addTable();
        for ($row = 1; $row <= $total_rows; $row++) {
            $table->addRow();
            for ( $cell = 1; $cell <= $total_cells; $cell++ ) {
                $table->addCell()->addText( "{$cell_text}" );
            }
        }
        
        $objWriter = PhpWordIOFactory::createWriter( $phpWord, 'Word2007' );
        $objWriter->save( 'Addresses.docx' );

        return true;
    }
}
