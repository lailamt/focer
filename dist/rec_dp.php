<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Csv;

class RecargaDP
{
    function Eckhardt()
    {
        require '../vendor/autoload.php';
        $ext = strtolower(substr($_FILES['uploadFile']['name'], -4));
        date_default_timezone_set('America/Sao_Paulo');
        $new_name = mt_rand() . $ext;
        $dir = 'uploads/';
        move_uploaded_file($_FILES['uploadFile']['tmp_name'], $dir . $new_name);
        $parA = ($_POST['formParametroA']);
        $parBFI = ($_POST['formParametroBFI']);
        $area = ($_POST['formAreaDrenagem']);
        $spreadsheet = IOFactory::load($dir . $new_name);
        $sheet = $spreadsheet->getActivesheet();
        $sheet->getStyle('B')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('C')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('D')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('E')->getNumberFormat()->setFormatCode('0.00');

        $sheet->setCellValue('A1', 'Tempo');
        $sheet->setCellValue('B1', 'Esc. Total (m³/s)');
        $sheet->setCellValue('C1', 'Esc. Base (m³/s)');
        $sheet->setCellValue('D1', 'Recarga (mm/dia)');
        $sheet->setCellValue('E1', 'Recarga anual (mm/ano)');
        $sheet->setCellValue('F1', 'Área de drenagem (km²)');
        $n = 1;
        $nn = 2;
        $lines = 2;
        while ($lines <= $sheet->getHighestRow()) {
            $rowNum = $n + 1;
            $rowNumMinus = $nn - 1;
            $rowNumMax = $sheet->getHighestRow();
            $row['Tempo'] = $sheet->getCell("A" . $lines)->getValue();
            $row['Esc. Total (m³/s)'] = $sheet->getCell("B" . $lines)->getValue();
            
            $sheet->setCellValue('A' . $rowNum, $row['Tempo']);
            $sheet->setCellValue('B' . $rowNum, $row['Esc. Total (m³/s)']);
            $sheet->setCellValue('C2', '=B2');
            $sheet->setCellValue('C' . $rowNum, '=IF((((1-' . $parBFI . ')*' . $parA . '*' . 'C' . $rowNumMinus . '+(1-' . $parA . ')*' . $parBFI . '*' . 'B' . $rowNum . ')' . '/(1-' . $parA . '*' . $parBFI . '))>' . 'B' . $rowNum . ',' . 'B' . $rowNum . ',((1-' . $parBFI . ')*' . $parA . '*' . 'C' . $rowNumMinus . '+(1-' . $parA . ')*' . $parBFI . '*' . 'B' . $rowNum . ')' . '/(1-' . $parA . '*' . $parBFI . '))');
            $sheet->setCellValue('D' . $rowNum, '=(' . 'C' . $rowNum . '/(1000000*' . 'F2' . '))' . '*1000*(60*60*24)');
            $sheet->setCellValue('E2', '=AVERAGE(D2:' . $rowNumMax . ')*366');
            $sheet->setCellValue('F2', $area);
            $n++;
            $nn++;
            $lines++;
        }
        $filename = 'FOCER_'  . $new_name;
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer = new Csv($spreadsheet);
        $writer->setDelimiter(';');
        $writer->setEnclosure('"');
        $writer->setUseBOM(true);
        $writer->save('php://output');
        unlink($dir . $new_name);
        exit();
    }
    function LyneHollick()
    {
        require '../vendor/autoload.php';
        $ext = strtolower(substr($_FILES['uploadFile']['name'], -4));
        date_default_timezone_set('America/Sao_Paulo');
        $new_name = mt_rand() . $ext;
        $dir = 'uploads/';
        move_uploaded_file($_FILES['uploadFile']['tmp_name'], $dir . $new_name);
        $parA = ($_POST['formParametroA']);
        $area = ($_POST['formAreaDrenagem']);
        $spreadsheet = IOFactory::load($dir . $new_name);
        $sheet = $spreadsheet->getActivesheet();
        $sheet->getStyle('B')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('C')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('D')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('E')->getNumberFormat()->setFormatCode('0.00');

        $sheet->setCellValue('A1', 'Tempo');
        $sheet->setCellValue('B1', 'Esc. Total (m³/s)');
        $sheet->setCellValue('C1', 'Esc. Base (m³/s)');
        $sheet->setCellValue('D1', 'Recarga (mm/dia)');
        $sheet->setCellValue('E1', 'Recarga anual (mm/ano)');
        $sheet->setCellValue('F1', 'Área de drenagem (km²)');
        $n = 1;
        $nn = 2;
        $lines = 2;
        while ($lines <= $sheet->getHighestRow()) {
            $rowNum = $n + 1;
            $rowNumMinus = $nn - 1;
            $rowNumMax = $sheet->getHighestRow();
            $row['Tempo'] = $sheet->getCell("A" . $lines)->getValue();
            $row['Esc. Total (m³/s)'] = $sheet->getCell("B" . $lines)->getValue();

            $sheet->setCellValue('A' . $rowNum, $row['Tempo']);
            $sheet->setCellValue('B' . $rowNum, $row['Esc. Total (m³/s)']);
            $sheet->setCellValue('C2', '=B2');
            $sheet->setCellValue('C' . $rowNum, '=IF((' . $parA . '*' . 'C' . $rowNumMinus . '+((1-' . $parA . ')/2)*(' . 'B' . $rowNum . '+' . 'B' . $rowNumMinus . '))>' . 'B' . $rowNum . ',' . 'B' . $rowNum . ',' . $parA . '*' . 'C' . $rowNumMinus . '+((1-' . $parA . ')/2)*(' . 'B' . $rowNum . '+' . 'B' . $rowNumMinus . '))');
            $sheet->setCellValue('D' . $rowNum, '=(' . 'C' . $rowNum . '/(1000000*' . 'F2' . '))' . '*1000*(60*60*24)');
            $sheet->setCellValue('E2', '=AVERAGE(D2:' . $rowNumMax . ')*366');
            $sheet->setCellValue('F2', $area);
            $n++;
            $nn++;
            $lines++;
        }
        $filename = 'FOCER_'  . $new_name . '.csv';
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer = new Csv($spreadsheet);
        $writer->setDelimiter(';');
        $writer->setEnclosure('"');
        $writer->setUseBOM(true);
        $writer->save('php://output');
        unlink($dir . $new_name);
        exit();
    }
    function ChapmanMaxwell()
    {
        require '../vendor/autoload.php';
        $ext = strtolower(substr($_FILES['uploadFile']['name'], -4));
        date_default_timezone_set('America/Sao_Paulo');
        $new_name = mt_rand() . $ext;
        $dir = 'uploads/';
        move_uploaded_file($_FILES['uploadFile']['tmp_name'], $dir . $new_name);
        $parA = ($_POST['formParametroA']);
        $area = ($_POST['formAreaDrenagem']);
        $spreadsheet = IOFactory::load($dir . $new_name);
        $sheet = $spreadsheet->getActivesheet();
        $sheet->getStyle('B')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('C')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('D')->getNumberFormat()->setFormatCode('0.00');
        $sheet->getStyle('E')->getNumberFormat()->setFormatCode('0.00');

        $sheet->setCellValue('A1', 'Tempo');
        $sheet->setCellValue('B1', 'Esc. Total (m³/s)');
        $sheet->setCellValue('C1', 'Esc. Base (m³/s)');
        $sheet->setCellValue('D1', 'Recarga (mm/dia)');
        $sheet->setCellValue('E1', 'Recarga anual (mm/ano)');
        $sheet->setCellValue('F1', 'Área de drenagem (km²)');
        $n = 1;
        $nn = 2;
        $lines = 2;
        while ($lines <= $sheet->getHighestRow()) {
            $rowNum = $n + 1;
            $rowNumMinus = $nn - 1;
            $rowNumMax = $sheet->getHighestRow();
            $row['Tempo'] = $sheet->getCell("A" . $lines)->getValue();
            $row['Esc. Total (m³/s)'] = $sheet->getCell("B" . $lines)->getValue();

            $sheet->setCellValue('A' . $rowNum, $row['Tempo']);
            $sheet->setCellValue('B' . $rowNum, $row['Esc. Total (m³/s)']);
            $sheet->setCellValue('C2', '=B2');
            $sheet->setCellValue('C' . $rowNum, '=IF((' . $parA . '/(2-' . $parA . ')*' . 'C' . $rowNumMinus . ')+((1-' . $parA . ')/(2-' . $parA . '))*' . 'B' . $rowNum.'>' . 'B' . $rowNum . ',' . 'B' . $rowNum . ',(' . $parA . '/(2-' . $parA . ')*' . 'C' . $rowNumMinus . ')+((1-' . $parA . ')/(2-' . $parA . '))*' . 'B' . $rowNum.')');
            $sheet->setCellValue('D' . $rowNum, '=(' . 'C' . $rowNum . '/(1000000*' . 'F2' . '))' . '*1000*(60*60*24)');
            $sheet->setCellValue('E2', '=AVERAGE(D2:' . $rowNumMax . ')*366');
            $sheet->setCellValue('F2', $area);
            $n++;
            $nn++;
            $lines++;
        }
        $filename = 'FOCER_'  . $new_name . '.csv';
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="' . $filename . '"');
        header('Cache-Control: max-age=0');
        $writer = new Csv($spreadsheet);
        $writer->setDelimiter(';');
        $writer->setEnclosure('"');
        $writer->setUseBOM(true);
        $writer->save('php://output');
        unlink($dir . $new_name);
        exit();
    }
}
