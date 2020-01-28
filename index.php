<?php


require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$spreadsheet->getDefaultStyle()->getFont()->setName('Calibri');
$spreadsheet->getDefaultStyle()->getFont()->setSize(8);
$sheet = $spreadsheet->getActiveSheet();
$sheet->getStyle('B')->getAlignment()->setHorizontal('center');
$sheet->mergeCells("B".(1).":C".(1));
$sheet->setCellValue('B1', 'SEGURO SOCIAL UNIVERSITARIO');
$sheet->mergeCells("B".(2).":C".(2));
$sheet->setCellValue('B2', 'ORURO-BOLIVIA');
$sheet->mergeCells("B".(3).":P".(3));
$sheet->setCellValue('B3', '1250101 FARMACIA DE CENTROS SANITARIOS');
$sheet->mergeCells("B".(4).":P".(4));
$sheet->setCellValue('B4', 'MEDICAMENTOS: MEDICAMENTOS');
$sheet->mergeCells("B".(5).":P".(5));
$sheet->setCellValue('B5', 'KARDEZ FISIO-VALORADO');
$sheet->mergeCells("B".(6).":P".(6));
$sheet->setCellValue('B6', 'AL 31 DE DICIEMBRE DE 2017');
$sheet->mergeCells("B".(7).":P".(7));
$sheet->setCellValue('B7', '(Expresado en Bolivianos)');
$sheet->setCellValue('B11', 'CODIGO:');
$sheet->setCellValue('C11', '02-02');
$sheet->setCellValue('N11', 'UNIDAD:');
$sheet->setCellValue('O11', 'COMP.');
$sheet->setCellValue('B12', 'DESCRIPCION:');
$sheet->setCellValue('C12', 'AMOXICILINA');
$sheet->setCellValue('N12', 'CONCEN:');
$sheet->setCellValue('O12', '1 GR');
$sheet->mergeCells("B14:B15");
$sheet->setCellValue('B14', 'FECHA');
$sheet->mergeCells("C14:C15");
$sheet->setCellValue('C14', 'No E/S');
$sheet->mergeCells("D14:D15");
$sheet->setCellValue('D14', 'DESCRIPCION');

$sheet->setCellValue('D15', 'CPRO_CODIGO');
$sheet->setCellValue('E15', 'CPRO_DESCRIPCION');
$sheet->setCellValue('F15', 'CPRO_CONCENT');
$sheet->setCellValue('G15', 'CPROU_DESCRIPCION');
$sheet->setCellValue('I15', 'PRODUCTO');
$sheet->mergeCells("J14:L14");
$sheet->setCellValue('J14', 'CANTIDAD');
$sheet->setCellValue('J15', 'INGRESO');
$sheet->setCellValue('K15', 'EGRESO');
$sheet->setCellValue('L15', 'SALDO');
$sheet->mergeCells("M14:M15");
$sheet->setCellValue('M14', 'PRECIO UNITARIO');
$sheet->mergeCells("N14:P14");
$sheet->setCellValue('N14', 'VALORADO');
$sheet->setCellValue('N15', 'INGRESO');
$sheet->setCellValue('O15', 'EGRESO');
$sheet->setCellValue('P15', 'SALDO');

$writer = new Xlsx($spreadsheet);
$writer->save('a.xlsx');

header("Content-disposition: attachment; filename=a.xlsx");
header("Content-type: application/xlsx");
readfile("a.xlsx");
