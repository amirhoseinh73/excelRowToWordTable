<?php

namespace App\Controllers;

use phpDocumentor\Reflection\PseudoTypes\True_;
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
        $section = $phpWord->addSection();

        // $table = $section->addTable();

        foreach( $rows_and_columns as $index => $row ) {

            if ( $index === 1 ) continue;
            $sender = "هابینو";
            $address = "آدرس: ";
            $post_code = "کد پستی: ";
            $mobile = "تلفن: ";
            $fullname = "نام و نام خانوادگی: ";
            $order_status = "وضعیت سفارش: ";
            $order_date = "تاریخ سفارش: ";
            $post_method = "نوع ارسال: ";

            foreach( $row as $idx => $column ) {
                
                // get address
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: city" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: state" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: street address (full)" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: city" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: state" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: street address (full)"
                ) $address .= $column . " - ";

                // get post code
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: zip code" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: zip code"
                ) $post_code .= $column;

                // get mobile
                if (
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "billing: phone number" ||
                    trim( strtolower( $rows_and_columns[ 1 ][ $idx ] ) ) === "shipping: phone number"
                ) $mobile .= $column;

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

            // $table->addRow();
            // $table->addCell()->addText( "{$sender}" );
            // $table->addCell()->addText( "{$address}" );
            // $table->addCell()->addText( "{$post_code}" );
            // $table->addCell()->addText( "{$mobile}" );
            // $table->addCell()->addText( "{$fullname}" );
            // $table->addCell()->addText( "{$order_status}" );
            // $table->addCell()->addText( "{$order_date}" );
            // $table->addCell()->addText( "{$post_method}" );

            $section->addText( "{$sender}" );
            $section->addText( "{$address}" );
            $section->addText( "{$post_code}" );
            $section->addText( "{$mobile}" );
            $section->addText( "{$fullname}" );
            $section->addText( "{$order_status}" );
            $section->addText( "{$order_date}" );
            $section->addText( "{$post_method}" );

        }
        
        $phpWord->setDefaultFontSize(18);
        // $phpWord->colsNum( 2 );
        $phpWord->setDefaultFontName('Arial');

        $objWriter = PhpWordIOFactory::createWriter( $phpWord, 'Word2007' );
        $file_word_name = "Addresses.docx";
        $objWriter->save( FCPATH . $file_word_name );

        echo "<a href='" . base_url( $file_word_name ) . "' id='download' >donwload file</a>
            <script>
                document.getElementById( 'download' ).click();
                setTimeout( () => { window.location.href='" . base_url( '/' ) . "' }, 500 );
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
