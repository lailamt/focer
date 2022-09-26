<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$radiobtn = $_GET['metodo'];
if ($radiobtn == 'metod1_ana') {

    /**
     * consulta os dados da ANA e retorna o conteúdo numa string formatado como xml 
     * @return string
     */
    function get_xml_from_ana()
    {
        $curl = curl_init();

        $codEstacao = ($_GET['codEstacao']);
        $dataInicio = ($_GET['dataInicio']);
        $dataFim = ($_GET['dataFim']);
        $nivelConsistencia = $_GET['nivelConsistencia'];
        $tipoDados = '3';

        curl_setopt_array($curl, [
            CURLOPT_URL => "http://telemetriaws1.ana.gov.br//ServiceANA.asmx/HidroSerieHistorica?codEstacao={$codEstacao}&dataInicio={$dataInicio}&dataFim={$dataFim}&tipoDados=3&nivelConsistencia={$nivelConsistencia}",
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "",
            CURLOPT_MAXREDIRS => 10,
            CURLOPT_TIMEOUT => 30,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => "GET",
            CURLOPT_POSTFIELDS => "",
        ]);

        $response = curl_exec($curl);
        $err = curl_error($curl);

        curl_close($curl);

        if ($err) {
            return false;
        }
        return (string) $response;
    }

    /**
     * Recebe o xml bruto, filtra apenas com os dados e retorna num array associativo
     * @param string $dataAsXmlString
     * @return array
     */
    function xml_to_assoc_array(string $dataAsXmlString)
    {
        $filters = [
            'EstacaoCodigo',
            'NivelConsistencia',
            'DataHora',
            'Vazao01',
            'Vazao02',
            'Vazao03',
            'Vazao04',
            'Vazao05',
            'Vazao06',
            'Vazao07',
            'Vazao08',
            'Vazao09',
            'Vazao10',
            'Vazao11',
            'Vazao12',
            'Vazao13',
            'Vazao14',
            'Vazao15',
            'Vazao16',
            'Vazao17',
            'Vazao18',
            'Vazao19',
            'Vazao20',
            'Vazao21',
            'Vazao22',
            'Vazao23',
            'Vazao24',
            'Vazao25',
            'Vazao26',
            'Vazao27',
            'Vazao28',
            'Vazao29',
            'Vazao30',
            'Vazao31',

        ];
        $regexFilter = implode('|', $filters);
        $array = [];

        $patern = '/\<SerieHistorica diffgr\:id\=\"SerieHistorica[0-9]+\" msdata\:rowOrder\=\"[0-9]+\"\>([^#]*?)\<\/SerieHistorica\>/';
        preg_match_all($patern, $dataAsXmlString, $filtered);
        $serieHistoricaList = $filtered[1];
        // $serieHistorica = $serieHistoricaList[0];
        foreach ($serieHistoricaList as $serieHistorica) {
            $patern = "/\<($regexFilter)?\>(.*?)\<\/.*\>/";
            preg_match_all($patern, $serieHistorica, $filtered);
            $titleList = $filtered[1];
            $valueList = $filtered[2];
            // var_dump($valueList);die;
            $serieHistoricaAsArray = [];
            for ($i = 0; $i < count($titleList); $i++) {
                $serieHistoricaAsArray[$titleList[$i]] = $valueList[$i];
            }
            $array[] = $serieHistoricaAsArray;
        }
        return $array;
    }

    /**
     * Recebe o array contendo os dados onde o primeiro indice é a header e o restante são os dados
     * @param array $data
     * @return bool
     */
    function export_to_xls(array $data)
    {

        $parA = ($_GET['formParametroA']);
        $parBFI = ($_GET['formParametroBFI']);
        $area = ($_GET['formAreaDrenagem']);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray($data);
        $sheet->setCellValue('AH1', 'Vazao31');
        $sheet->setCellValue('AI1', 'VazaoBase01');
        $sheet->setCellValue('AJ1', 'VazaoBase02');
        $sheet->setCellValue('AK1', 'VazaoBase03');
        $sheet->setCellValue('AL1', 'VazaoBase04');
        $sheet->setCellValue('AM1', 'VazaoBase05');
        $sheet->setCellValue('AN1', 'VazaoBase06');
        $sheet->setCellValue('AO1', 'VazaoBase07');
        $sheet->setCellValue('AP1', 'VazaoBase08');
        $sheet->setCellValue('AQ1', 'VazaoBase09');
        $sheet->setCellValue('AR1', 'VazaoBase10');
        $sheet->setCellValue('AS1', 'VazaoBase11');
        $sheet->setCellValue('AT1', 'VazaoBase12');
        $sheet->setCellValue('AU1', 'VazaoBase13');
        $sheet->setCellValue('AV1', 'VazaoBase14');
        $sheet->setCellValue('AW1', 'VazaoBase15');
        $sheet->setCellValue('AX1', 'VazaoBase16');
        $sheet->setCellValue('AY1', 'VazaoBase17');
        $sheet->setCellValue('AZ1', 'VazaoBase18');
        $sheet->setCellValue('BA1', 'VazaoBase19');
        $sheet->setCellValue('BB1', 'VazaoBase20');
        $sheet->setCellValue('BC1', 'VazaoBase21');
        $sheet->setCellValue('BD1', 'VazaoBase22');
        $sheet->setCellValue('BE1', 'VazaoBase23');
        $sheet->setCellValue('BF1', 'VazaoBase24');
        $sheet->setCellValue('BG1', 'VazaoBase25');
        $sheet->setCellValue('BH1', 'VazaoBase26');
        $sheet->setCellValue('BI1', 'VazaoBase27');
        $sheet->setCellValue('BJ1', 'VazaoBase28');
        $sheet->setCellValue('BK1', 'VazaoBase29');
        $sheet->setCellValue('BL1', 'VazaoBase30');
        $sheet->setCellValue('BM1', 'VazaoBase31');
        $sheet->setCellValue('BN1', 'RecargaDia01');
        $sheet->setCellValue('BO1', 'RecargaDia02');
        $sheet->setCellValue('BP1', 'RecargaDia03');
        $sheet->setCellValue('BQ1', 'RecargaDia04');
        $sheet->setCellValue('BR1', 'RecargaDia05');
        $sheet->setCellValue('BS1', 'RecargaDia06');
        $sheet->setCellValue('BT1', 'RecargaDia07');
        $sheet->setCellValue('BU1', 'RecargaDia08');
        $sheet->setCellValue('BV1', 'RecargaDia09');
        $sheet->setCellValue('BW1', 'RecargaDia10');
        $sheet->setCellValue('BX1', 'RecargaDia11');
        $sheet->setCellValue('BY1', 'RecargaDia12');
        $sheet->setCellValue('BZ1', 'RecargaDia13');
        $sheet->setCellValue('CA1', 'RecargaDia14');
        $sheet->setCellValue('CB1', 'RecargaDia15');
        $sheet->setCellValue('CC1', 'RecargaDia16');
        $sheet->setCellValue('CD1', 'RecargaDia17');
        $sheet->setCellValue('CE1', 'RecargaDia18');
        $sheet->setCellValue('CF1', 'RecargaDia19');
        $sheet->setCellValue('CG1', 'RecargaDia20');
        $sheet->setCellValue('CH1', 'RecargaDia21');
        $sheet->setCellValue('CI1', 'RecargaDia22');
        $sheet->setCellValue('CJ1', 'RecargaDia23');
        $sheet->setCellValue('CK1', 'RecargaDia24');
        $sheet->setCellValue('CL1', 'RecargaDia25');
        $sheet->setCellValue('CM1', 'RecargaDia26');
        $sheet->setCellValue('CN1', 'RecargaDia27');
        $sheet->setCellValue('CO1', 'RecargaDia28');
        $sheet->setCellValue('CP1', 'RecargaDia29');
        $sheet->setCellValue('CQ1', 'RecargaDia30');
        $sheet->setCellValue('CR1', 'RecargaDia31');

        $sheet->getStyle('D:CR')->getNumberFormat()->setFormatCode('0.00');

        $n = 1;
        $nn = 2;

        for ($i = 1; $i < count($data); $i++) {
            $rowNum = $n + 1;
            $rowNumMinus = $nn - 1;

            /*Vazão de base pelo método dos filtros númericos de Eckhardt (2005)*/
            /*VazaoBase01-Início da série histórica*/
            $sheet->setCellValue('AI2', '=IF(D2="","",D2)');
            /*VazaoBase01 - Após período inicial [DataInicio]*/
            $sheet->setCellValue('AI' . $rowNum, '=IF((AI' . $rowNumMinus . ':BM' . $rowNumMinus . ')="",IF(D' . $rowNum . '="","",D' . $rowNum . '),IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNumMinus . ':BM' . $rowNumMinus . '<>""),$AI$' . $rowNumMinus . ':BM' . $rowNumMinus . '))+(1-' . $parA . ')*' . $parBFI . '*D' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>D' . $rowNum . ', IF(D' . $rowNum . '="","",D' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNumMinus . ':BM' . $rowNumMinus . '<>""),$AI$' . $rowNumMinus . ':BM' . $rowNumMinus . '))+(1-' . $parA . ')*' . $parBFI . '*D' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')))');
            /*VazaoBase02*/
            $sheet->setCellValue('AJ' . $rowNum, '=IF(E' . $rowNum . '="","",IF(AI' . $rowNum . '="",E' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/(AI' . $rowNum . '<>""),AI' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*E' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>E' . $rowNum . ', IF(E' . $rowNum . '="","",E' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/(AI' . $rowNum . '<>""),AI' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*E' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase03*/
            $sheet->setCellValue('AK' . $rowNum, '=IF(F' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AJ' . $rowNum  . ')<=0,F' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AJ' . $rowNum . '<>""),$AI$' . $rowNum . ':AJ' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*F' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>F' . $rowNum . ', IF(F' . $rowNum . '="","",F' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AJ' . $rowNum . '<>""),$AI$' . $rowNum . ':AJ' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*F' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase04*/
            $sheet->setCellValue('AL' . $rowNum, '=IF(G' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AK' . $rowNum  . ')<=0,G' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AK' . $rowNum . '<>""),$AI$' . $rowNum . ':AK' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*G' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>G' . $rowNum . ', IF(G' . $rowNum . '="","",G' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AK' . $rowNum . '<>""),$AI$' . $rowNum . ':AK' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*G' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase05*/
            $sheet->setCellValue('AM' . $rowNum, '=IF(H' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AL' . $rowNum  . ')<=0,H' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AL' . $rowNum . '<>""),$AI$' . $rowNum . ':AL' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*H' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>H' . $rowNum . ', IF(H' . $rowNum . '="","",H' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AL' . $rowNum . '<>""),$AI$' . $rowNum . ':AL' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*H' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase06*/
            $sheet->setCellVAlue('AN' . $rowNum, '=IF(I' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AM' . $rowNum  . ')<=0,I' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AM' . $rowNum . '<>""),$AI$' . $rowNum . ':AM' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*I' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>I' . $rowNum . ', IF(I' . $rowNum . '="","",I' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AM' . $rowNum . '<>""),$AI$' . $rowNum . ':AM' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*I' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase07*/
            $sheet->setCellVAlue('AO' . $rowNum, '=IF(J' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AN' . $rowNum  . ')<=0,J' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AN' . $rowNum . '<>""),$AI$' . $rowNum . ':AN' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*J' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>J' . $rowNum . ', IF(J' . $rowNum . '="","",J' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AN' . $rowNum . '<>""),$AI$' . $rowNum . ':AN' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*J' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase08*/
            $sheet->setCellVAlue('AP' . $rowNum, '=IF(K' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AO' . $rowNum  . ')<=0,K' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AO' . $rowNum . '<>""),$AI$' . $rowNum . ':AO' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*K' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>K' . $rowNum . ', IF(K' . $rowNum . '="","",K' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AO' . $rowNum . '<>""),$AI$' . $rowNum . ':AO' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*K' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase09*/
            $sheet->setCellVAlue('AQ' . $rowNum, '=IF(L' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AP' . $rowNum  . ')<=0,L' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AP' . $rowNum . '<>""),$AI$' . $rowNum . ':AP' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*L' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>L' . $rowNum . ', IF(L' . $rowNum . '="","",L' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AP' . $rowNum . '<>""),$AI$' . $rowNum . ':AP' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*L' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase10*/
            $sheet->setCellVAlue('AR' . $rowNum, '=IF(M' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AQ' . $rowNum  . ')<=0,M' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AQ' . $rowNum . '<>""),$AI$' . $rowNum . ':AQ' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*M' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>M' . $rowNum . ', IF(M' . $rowNum . '="","",M' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AQ' . $rowNum . '<>""),$AI$' . $rowNum . ':AQ' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*M' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase11*/
            $sheet->setCellVAlue('AS' . $rowNum, '=IF(N' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AR' . $rowNum  . ')<=0,N' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AR' . $rowNum . '<>""),$AI$' . $rowNum . ':AR' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*N' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>N' . $rowNum . ', IF(N' . $rowNum . '="","",N' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AR' . $rowNum . '<>""),$AI$' . $rowNum . ':AR' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*N' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase12*/
            $sheet->setCellVAlue('AT' . $rowNum, '=IF(O' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AS' . $rowNum  . ')<=0,O' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AS' . $rowNum . '<>""),$AI$' . $rowNum . ':AS' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*O' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>O' . $rowNum . ', IF(O' . $rowNum . '="","",O' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AS' . $rowNum . '<>""),$AI$' . $rowNum . ':AS' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*O' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase13*/
            $sheet->setCellVAlue('AU' . $rowNum, '=IF(P' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AT' . $rowNum  . ')<=0,P' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AT' . $rowNum . '<>""),$AI$' . $rowNum . ':AT' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*P' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>P' . $rowNum . ', IF(P' . $rowNum . '="","",P' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AT' . $rowNum . '<>""),$AI$' . $rowNum . ':AT' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*P' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase14*/
            $sheet->setCellVAlue('AV' . $rowNum, '=IF(Q' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AU' . $rowNum  . ')<=0,Q' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AU' . $rowNum . '<>""),$AI$' . $rowNum . ':AU' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*Q' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>Q' . $rowNum . ', IF(Q' . $rowNum . '="","",Q' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AU' . $rowNum . '<>""),$AI$' . $rowNum . ':AU' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*Q' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase15*/
            $sheet->setCellVAlue('AW' . $rowNum, '=IF(R' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AV' . $rowNum  . ')<=0,R' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AV' . $rowNum . '<>""),$AI$' . $rowNum . ':AV' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*R' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>R' . $rowNum . ', IF(R' . $rowNum . '="","",R' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AV' . $rowNum . '<>""),$AI$' . $rowNum . ':AV' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*R' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase16*/
            $sheet->setCellVAlue('AX' . $rowNum, '=IF(S' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AW' . $rowNum  . ')<=0,S' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AW' . $rowNum . '<>""),$AI$' . $rowNum . ':AW' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*S' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>S' . $rowNum . ', IF(S' . $rowNum . '="","",S' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AW' . $rowNum . '<>""),$AI$' . $rowNum . ':AW' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*S' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase17*/
            $sheet->setCellVAlue('AY' . $rowNum, '=IF(T' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AX' . $rowNum  . ')<=0,T' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AX' . $rowNum . '<>""),$AI$' . $rowNum . ':AX' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*T' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>T' . $rowNum . ', IF(T' . $rowNum . '="","",T' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AX' . $rowNum . '<>""),$AI$' . $rowNum . ':AX' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*T' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase18*/
            $sheet->setCellVAlue('AZ' . $rowNum, '=IF(U' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AY' . $rowNum  . ')<=0,U' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AY' . $rowNum . '<>""),$AI$' . $rowNum . ':AY' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*U' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>U' . $rowNum . ', IF(U' . $rowNum . '="","",U' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AY' . $rowNum . '<>""),$AI$' . $rowNum . ':AY' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*U' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase19*/
            $sheet->setCellVAlue('BA' . $rowNum, '=IF(V' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AZ' . $rowNum  . ')<=0,V' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AZ' . $rowNum . '<>""),$AI$' . $rowNum . ':AZ' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*V' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>V' . $rowNum . ', IF(V' . $rowNum . '="","",V' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AZ' . $rowNum . '<>""),$AI$' . $rowNum . ':AZ' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*V' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase20*/
            $sheet->setCellVAlue('BB' . $rowNum, '=IF(W' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BA' . $rowNum  . ')<=0,W' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BA' . $rowNum . '<>""),$AI$' . $rowNum . ':BA' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*W' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>W' . $rowNum . ', IF(W' . $rowNum . '="","",W' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BA' . $rowNum . '<>""),$AI$' . $rowNum . ':BA' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*W' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase21*/
            $sheet->setCellVAlue('BC' . $rowNum, '=IF(X' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BB' . $rowNum  . ')<=0,X' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BB' . $rowNum . '<>""),$AI$' . $rowNum . ':BB' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*X' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>X' . $rowNum . ', IF(X' . $rowNum . '="","",X' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BB' . $rowNum . '<>""),$AI$' . $rowNum . ':BB' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*X' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase22*/
            $sheet->setCellVAlue('BD' . $rowNum, '=IF(Y' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BC' . $rowNum  . ')<=0,Y' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BC' . $rowNum . '<>""),$AI$' . $rowNum . ':BC' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*Y' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>Y' . $rowNum . ', IF(Y' . $rowNum . '="","",Y' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BC' . $rowNum . '<>""),$AI$' . $rowNum . ':BC' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*Y' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase23*/
            $sheet->setCellVAlue('BE' . $rowNum, '=IF(Z' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BD' . $rowNum  . ')<=0,Z' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BD' . $rowNum . '<>""),$AI$' . $rowNum . ':BD' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*Z' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>Z' . $rowNum . ', IF(Z' . $rowNum . '="","",Z' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BD' . $rowNum . '<>""),$AI$' . $rowNum . ':BD' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*Z' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase24*/
            $sheet->setCellVAlue('BF' . $rowNum, '=IF(AA' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BE' . $rowNum  . ')<=0,AA' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BE' . $rowNum . '<>""),$AI$' . $rowNum . ':BE' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*AA' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>AA' . $rowNum . ', IF(AA' . $rowNum . '="","",AA' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BE' . $rowNum . '<>""),$AI$' . $rowNum . ':BE' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*AA' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase25*/
            $sheet->setCellVAlue('BG' . $rowNum, '=IF(AB' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BF' . $rowNum  . ')<=0,AB' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BF' . $rowNum . '<>""),$AI$' . $rowNum . ':BF' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*AB' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>AB' . $rowNum . ', IF(AB' . $rowNum . '="","",AB' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BF' . $rowNum . '<>""),$AI$' . $rowNum . ':BF' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*AB' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase26*/
            $sheet->setCellVAlue('BH' . $rowNum, '=IF(AC' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BG' . $rowNum  . ')<=0,AC' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BG' . $rowNum . '<>""),$AI$' . $rowNum . ':BG' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*AC' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>AC' . $rowNum . ', IF(AC' . $rowNum . '="","",AC' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BG' . $rowNum . '<>""),$AI$' . $rowNum . ':BG' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*AC' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase27*/
            $sheet->setCellVAlue('BI' . $rowNum, '=IF(AD' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BH' . $rowNum  . ')<=0,AD' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BH' . $rowNum . '<>""),$AI$' . $rowNum . ':BH' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*AD' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>AD' . $rowNum . ', IF(AD' . $rowNum . '="","",AD' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BH' . $rowNum . '<>""),$AI$' . $rowNum . ':BH' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*AD' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase28*/
            $sheet->setCellVAlue('BJ' . $rowNum, '=IF(AE' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BI' . $rowNum  . ')<=0,AE' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BI' . $rowNum . '<>""),$AI$' . $rowNum . ':BI' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*AE' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>AE' . $rowNum . ', IF(AE' . $rowNum . '="","",AE' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BI' . $rowNum . '<>""),$AI$' . $rowNum . ':BI' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*AE' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase29*/
            $sheet->setCellVAlue('BK' . $rowNum, '=IF(AF' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BJ' . $rowNum  . ')<=0,AF' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BJ' . $rowNum . '<>""),$AI$' . $rowNum . ':BJ' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*AF' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>AF' . $rowNum . ', IF(AF' . $rowNum . '="","",AF' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BJ' . $rowNum . '<>""),$AI$' . $rowNum . ':BJ' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*AF' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase30*/
            $sheet->setCellVAlue('BL' . $rowNum, '=IF(AG' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BK' . $rowNum  . ')<=0,AG' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BK' . $rowNum . '<>""),$AI$' . $rowNum . ':BK' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*AG' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>AG' . $rowNum . ', IF(AG' . $rowNum . '="","",AG' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BK' . $rowNum . '<>""),$AI$' . $rowNum . ':BK' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*AG' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');
            /*VazaoBase31*/
            $sheet->setCellVAlue('BM' . $rowNum, '=IF(AH' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BL' . $rowNum  . ')<=0,AH' . $rowNum . ',IF(((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BL' . $rowNum . '<>""),$AI$' . $rowNum . ':BL' . $rowNum  . '))+(1-' . $parA . ')*' . $parBFI . '*AH' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . ')>AH' . $rowNum . ', IF(AH' . $rowNum . '="","",AH' . $rowNum . '),((1-' . $parBFI . ')*' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BL' . $rowNum . '<>""),$AI$' . $rowNum . ':BL' . $rowNum . '))+(1-' . $parA . ')*' . $parBFI . '*AH' . $rowNum . ')/(1-' . $parA . '*' . $parBFI . '))))');

            /*Recarga diária (mm/dia)*/
            /*RecargaDia01*/
            $sheet->setCellValue('BN' . $rowNum, '=IF(AI' . $rowNum . '="","",(AI' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia02*/
            $sheet->setCellValue('BO' . $rowNum, '=IF(AJ' . $rowNum . '="","",(AJ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia03*/
            $sheet->setCellValue('BP' . $rowNum, '=IF(AK' . $rowNum . '="","",(AK' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia04*/
            $sheet->setCellValue('BQ' . $rowNum, '=IF(AL' . $rowNum . '="","",(AL' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia05*/
            $sheet->setCellValue('BR' . $rowNum, '=IF(AM' . $rowNum . '="","",(AM' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia06*/
            $sheet->setCellValue('BS' . $rowNum, '=IF(AN' . $rowNum . '="","",(AN' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia07*/
            $sheet->setCellValue('BT' . $rowNum, '=IF(AO' . $rowNum . '="","",(AO' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia08*/
            $sheet->setCellValue('BU' . $rowNum, '=IF(AP' . $rowNum . '="","",(AP' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia09*/
            $sheet->setCellValue('BV' . $rowNum, '=IF(AQ' . $rowNum . '="","",(AQ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia10*/
            $sheet->setCellValue('BW' . $rowNum, '=IF(AR' . $rowNum . '="","",(AR' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia11*/
            $sheet->setCellValue('BX' . $rowNum, '=IF(AS' . $rowNum . '="","",(AS' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia12*/
            $sheet->setCellValue('BY' . $rowNum, '=IF(AT' . $rowNum . '="","",(AT' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia13*/
            $sheet->setCellValue('BZ' . $rowNum, '=IF(AU' . $rowNum . '="","",(AU' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia14*/
            $sheet->setCellValue('CA' . $rowNum, '=IF(AV' . $rowNum . '="","",(AV' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia15*/
            $sheet->setCellValue('CB' . $rowNum, '=IF(AW' . $rowNum . '="","",(AW' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia16*/
            $sheet->setCellValue('CC' . $rowNum, '=IF(AX' . $rowNum . '="","",(AX' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia17*/
            $sheet->setCellValue('CD' . $rowNum, '=IF(AY' . $rowNum . '="","",(AY' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia18*/
            $sheet->setCellValue('CE' . $rowNum, '=IF(AZ' . $rowNum . '="","",(AZ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia19*/
            $sheet->setCellValue('CF' . $rowNum, '=IF(BA' . $rowNum . '="","",(BA' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia20*/
            $sheet->setCellValue('CG' . $rowNum, '=IF(BB' . $rowNum . '="","",(BB' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia21*/
            $sheet->setCellValue('CH' . $rowNum, '=IF(BC' . $rowNum . '="","",(BC' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia22*/
            $sheet->setCellValue('CI' . $rowNum, '=IF(BD' . $rowNum . '="","",(BD' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia23*/
            $sheet->setCellValue('CJ' . $rowNum, '=IF(BE' . $rowNum . '="","",(BE' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia24*/
            $sheet->setCellValue('CK' . $rowNum, '=IF(BF' . $rowNum . '="","",(BF' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia25*/
            $sheet->setCellValue('CL' . $rowNum, '=IF(BG' . $rowNum . '="","",(BG' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia26*/
            $sheet->setCellValue('CM' . $rowNum, '=IF(BH' . $rowNum . '="","",(BH' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia27*/
            $sheet->setCellValue('CN' . $rowNum, '=IF(BI' . $rowNum . '="","",(BI' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia28*/
            $sheet->setCellValue('CO' . $rowNum, '=IF(BJ' . $rowNum . '="","",(BJ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia29*/
            $sheet->setCellValue('CP' . $rowNum, '=IF(BK' . $rowNum . '="","",(BK' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia30*/
            $sheet->setCellValue('CQ' . $rowNum, '=IF(BL' . $rowNum . '="","",(BL' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia30*/
            $sheet->setCellValue('CR' . $rowNum, '=IF(BM' . $rowNum . '="","",(BM' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');

            $n++;
            $nn++;
        }


        $writer = new Xlsx($spreadsheet);
        $writer->setPreCalculateFormulas(false);
        $codEstacao = ($_GET['codEstacao']);
        $fileName =  'FOCER-' . $codEstacao . '.Xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . urlencode($fileName) . '"');
        $writer->save('php://output');
    }

    /**
     * formata o array multidimensional associativo para um array multidimensional simples
     * onde o primeiro indice é um array com as keys e o restante são os dados
     * @param array $assoc_array
     * @return array
     */
    function format_array(array $assoc_array)
    {
        $array[] = array_keys($assoc_array[0]);
        for ($i = count($assoc_array) - 1; $i > 0; $i--) {
            $array[] = array_values($assoc_array[$i]);
        }
        return $array;
    }
    /**
     * método responsável por consultar os dados da ANA e exportar num arquivo xlsx
     */
    function export_data_from_ana()
    {
        $dataAsXmlString = get_xml_from_ana();
        if ($dataAsXmlString) {
            $dataAsAssocArray = xml_to_assoc_array($dataAsXmlString);
            $dataAsArray = format_array($dataAsAssocArray);
            export_to_xls($dataAsArray);
        }
    }
    export_data_from_ana();
} elseif ($radiobtn == 'metod2_ana') {

    /**
     * consulta os dados da ANA e retorna o conteúdo numa string formatado como xml 
     * @return string
     */
    function get_xml_from_ana()
    {
        $curl = curl_init();

        $codEstacao = ($_GET['codEstacao']);
        $dataInicio = ($_GET['dataInicio']);
        $dataFim = ($_GET['dataFim']);
        $nivelConsistencia = $_GET['nivelConsistencia'];
        $tipoDados = '3';

        curl_setopt_array($curl, [
            CURLOPT_URL => "http://telemetriaws1.ana.gov.br//ServiceANA.asmx/HidroSerieHistorica?codEstacao={$codEstacao}&dataInicio={$dataInicio}&dataFim={$dataFim}&tipoDados=3&nivelConsistencia={$nivelConsistencia}",
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "",
            CURLOPT_MAXREDIRS => 10,
            CURLOPT_TIMEOUT => 30,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => "GET",
            CURLOPT_POSTFIELDS => "",
        ]);

        $response = curl_exec($curl);
        $err = curl_error($curl);

        curl_close($curl);

        if ($err) {
            return false;
        }
        return (string) $response;
    }

    /**
     * Recebe o xml bruto, filtra apenas com os dados e retorna num array associativo
     * @param string $dataAsXmlString
     * @return array
     */
    function xml_to_assoc_array(string $dataAsXmlString)
    {
        $filters = [
            'EstacaoCodigo',
            'NivelConsistencia',
            'DataHora',
            'Vazao01',
            'Vazao02',
            'Vazao03',
            'Vazao04',
            'Vazao05',
            'Vazao06',
            'Vazao07',
            'Vazao08',
            'Vazao09',
            'Vazao10',
            'Vazao11',
            'Vazao12',
            'Vazao13',
            'Vazao14',
            'Vazao15',
            'Vazao16',
            'Vazao17',
            'Vazao18',
            'Vazao19',
            'Vazao20',
            'Vazao21',
            'Vazao22',
            'Vazao23',
            'Vazao24',
            'Vazao25',
            'Vazao26',
            'Vazao27',
            'Vazao28',
            'Vazao29',
            'Vazao30',
            'Vazao31',

        ];
        $regexFilter = implode('|', $filters);
        $array = [];

        $patern = '/\<SerieHistorica diffgr\:id\=\"SerieHistorica[0-9]+\" msdata\:rowOrder\=\"[0-9]+\"\>([^#]*?)\<\/SerieHistorica\>/';
        preg_match_all($patern, $dataAsXmlString, $filtered);
        $serieHistoricaList = $filtered[1];
        // $serieHistorica = $serieHistoricaList[0];
        foreach ($serieHistoricaList as $serieHistorica) {
            $patern = "/\<($regexFilter)?\>(.*?)\<\/.*\>/";
            preg_match_all($patern, $serieHistorica, $filtered);
            $titleList = $filtered[1];
            $valueList = $filtered[2];
            // var_dump($valueList);die;
            $serieHistoricaAsArray = [];
            for ($i = 0; $i < count($titleList); $i++) {
                $serieHistoricaAsArray[$titleList[$i]] = $valueList[$i];
            }
            $array[] = $serieHistoricaAsArray;
        }
        return $array;
    }

    /**
     * Recebe o array contendo os dados onde o primeiro indice é a header e o restante são os dados
     * @param array $data
     * @return bool
     */
    function export_to_xls(array $data)
    {

        $parA = ($_GET['formParametroA']);
        $parBFI = ($_GET['formParametroBFI']);
        $area = ($_GET['formAreaDrenagem']);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray($data);
        $sheet->setCellValue('AH1', 'Vazao31');
        $sheet->setCellValue('AI1', 'VazaoBase01');
        $sheet->setCellValue('AJ1', 'VazaoBase02');
        $sheet->setCellValue('AK1', 'VazaoBase03');
        $sheet->setCellValue('AL1', 'VazaoBase04');
        $sheet->setCellValue('AM1', 'VazaoBase05');
        $sheet->setCellValue('AN1', 'VazaoBase06');
        $sheet->setCellValue('AO1', 'VazaoBase07');
        $sheet->setCellValue('AP1', 'VazaoBase08');
        $sheet->setCellValue('AQ1', 'VazaoBase09');
        $sheet->setCellValue('AR1', 'VazaoBase10');
        $sheet->setCellValue('AS1', 'VazaoBase11');
        $sheet->setCellValue('AT1', 'VazaoBase12');
        $sheet->setCellValue('AU1', 'VazaoBase13');
        $sheet->setCellValue('AV1', 'VazaoBase14');
        $sheet->setCellValue('AW1', 'VazaoBase15');
        $sheet->setCellValue('AX1', 'VazaoBase16');
        $sheet->setCellValue('AY1', 'VazaoBase17');
        $sheet->setCellValue('AZ1', 'VazaoBase18');
        $sheet->setCellValue('BA1', 'VazaoBase19');
        $sheet->setCellValue('BB1', 'VazaoBase20');
        $sheet->setCellValue('BC1', 'VazaoBase21');
        $sheet->setCellValue('BD1', 'VazaoBase22');
        $sheet->setCellValue('BE1', 'VazaoBase23');
        $sheet->setCellValue('BF1', 'VazaoBase24');
        $sheet->setCellValue('BG1', 'VazaoBase25');
        $sheet->setCellValue('BH1', 'VazaoBase26');
        $sheet->setCellValue('BI1', 'VazaoBase27');
        $sheet->setCellValue('BJ1', 'VazaoBase28');
        $sheet->setCellValue('BK1', 'VazaoBase29');
        $sheet->setCellValue('BL1', 'VazaoBase30');
        $sheet->setCellValue('BM1', 'VazaoBase31');
        $sheet->setCellValue('BN1', 'RecargaDia01');
        $sheet->setCellValue('BO1', 'RecargaDia02');
        $sheet->setCellValue('BP1', 'RecargaDia03');
        $sheet->setCellValue('BQ1', 'RecargaDia04');
        $sheet->setCellValue('BR1', 'RecargaDia05');
        $sheet->setCellValue('BS1', 'RecargaDia06');
        $sheet->setCellValue('BT1', 'RecargaDia07');
        $sheet->setCellValue('BU1', 'RecargaDia08');
        $sheet->setCellValue('BV1', 'RecargaDia09');
        $sheet->setCellValue('BW1', 'RecargaDia10');
        $sheet->setCellValue('BX1', 'RecargaDia11');
        $sheet->setCellValue('BY1', 'RecargaDia12');
        $sheet->setCellValue('BZ1', 'RecargaDia13');
        $sheet->setCellValue('CA1', 'RecargaDia14');
        $sheet->setCellValue('CB1', 'RecargaDia15');
        $sheet->setCellValue('CC1', 'RecargaDia16');
        $sheet->setCellValue('CD1', 'RecargaDia17');
        $sheet->setCellValue('CE1', 'RecargaDia18');
        $sheet->setCellValue('CF1', 'RecargaDia19');
        $sheet->setCellValue('CG1', 'RecargaDia20');
        $sheet->setCellValue('CH1', 'RecargaDia21');
        $sheet->setCellValue('CI1', 'RecargaDia22');
        $sheet->setCellValue('CJ1', 'RecargaDia23');
        $sheet->setCellValue('CK1', 'RecargaDia24');
        $sheet->setCellValue('CL1', 'RecargaDia25');
        $sheet->setCellValue('CM1', 'RecargaDia26');
        $sheet->setCellValue('CN1', 'RecargaDia27');
        $sheet->setCellValue('CO1', 'RecargaDia28');
        $sheet->setCellValue('CP1', 'RecargaDia29');
        $sheet->setCellValue('CQ1', 'RecargaDia30');
        $sheet->setCellValue('CR1', 'RecargaDia31');

        $sheet->getStyle('D:CR')->getNumberFormat()->setFormatCode('0.00');

        $n = 1;
        $nn = 2;

        for ($i = 1; $i < count($data); $i++) {
            $rowNum = $n + 1;
            $rowNumMinus = $nn - 1;

            /*Vazão de base pelo método dos filtros númericos de Lyne & Hollick (1979)*/
            /*VazaoBase01-Início da série histórica*/
            $sheet->setCellValue('AI2', '=IF(D2="","",D2)');
            /*VazaoBase01 - Após período inicial [DataInicio]*/
            $sheet->setCellValue('AI' . $rowNum, '=IF((AI' . $rowNumMinus . ':BM' . $rowNumMinus . ')="",IF(D' . $rowNum . '="","",D' . $rowNum . '),IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNumMinus . ':BM' . $rowNumMinus . '<>""),$AI$' . $rowNumMinus . ':BM' . $rowNumMinus  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNumMinus . ':AH' . $rowNumMinus . '<>""),$D$' . $rowNumMinus . ':AH' . $rowNumMinus  . '))' . '+D' . $rowNum . '))>D' . $rowNum . ',IF(D' . $rowNum . '="","",D' . $rowNum . '),((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNumMinus . ':BM' . $rowNumMinus . '<>""),$AI$' . $rowNumMinus . ':BM' . $rowNumMinus  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNumMinus . ':AH' . $rowNumMinus . '<>""),$D$' . $rowNumMinus . ':AH' . $rowNumMinus  . '))' . '+D' . $rowNum . '))))');
            /*VazaoBase02*/
            $sheet->setCellValue('AJ' . $rowNum, '=IF(E' . $rowNum . '="","",IF(AI' . $rowNum . '="",E' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/(AI' . $rowNum . '<>""),AI' . $rowNum . ')))+((1-' . $parA . ')/2)*(E' . $rowNum . '+D' . $rowNum . '))>E' . $rowNum . ',IF(E' . $rowNum . '="","",E' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/(AI' . $rowNum . '<>""),AI' . $rowNum . ')))+((1-' . $parA . ')/2)*(E' . $rowNum . '+D' . $rowNum . '))))');
            /*VazaoBase03*/
            $sheet->setCellValue('AK' . $rowNum, '=IF(F' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AJ' . $rowNum  . ')<=0,F' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AJ' . $rowNum . '<>""),$AI$' . $rowNum . ':AJ' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':E' . $rowNum . '<>""),$D$' . $rowNum . ':E' . $rowNum  . '))' . '+F' . $rowNum . '))>F' . $rowNum . ',IF(F' . $rowNum . '="","",F' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AJ' . $rowNum . '<>""),$AI$' . $rowNum . ':AJ' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':E' . $rowNum . '<>""),$D$' . $rowNum . ':E' . $rowNum  . '))' . '+F' . $rowNum . '))))');
            /*VazaoBase04*/
            $sheet->setCellValue('AL' . $rowNum, '=IF(G' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AK' . $rowNum  . ')<=0,G' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AK' . $rowNum . '<>""),$AI$' . $rowNum . ':AK' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':F' . $rowNum . '<>""),$D$' . $rowNum . ':F' . $rowNum  . '))' . '+G' . $rowNum . '))>G' . $rowNum . ',IF(G' . $rowNum . '="","",G' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AK' . $rowNum . '<>""),$AI$' . $rowNum . ':AK' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':F' . $rowNum . '<>""),$D$' . $rowNum . ':F' . $rowNum  . '))' . '+G' . $rowNum . '))))');
            /*VazaoBase05*/
            $sheet->setCellValue('AM' . $rowNum, '=IF(H' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AL' . $rowNum  . ')<=0,H' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AL' . $rowNum . '<>""),$AI$' . $rowNum . ':AL' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':G' . $rowNum . '<>""),$D$' . $rowNum . ':G' . $rowNum  . '))' . '+H' . $rowNum . '))>H' . $rowNum . ',IF(H' . $rowNum . '="","",H' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AL' . $rowNum . '<>""),$AI$' . $rowNum . ':AL' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':G' . $rowNum . '<>""),$D$' . $rowNum . ':G' . $rowNum  . '))' . '+H' . $rowNum . '))))');
            /*VazaoBase06*/
            $sheet->setCellVAlue('AN' . $rowNum, '=IF(I' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AM' . $rowNum  . ')<=0,I' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AM' . $rowNum . '<>""),$AI$' . $rowNum . ':AM' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':H' . $rowNum . '<>""),$D$' . $rowNum . ':H' . $rowNum  . '))' . '+I' . $rowNum . '))>I' . $rowNum . ',IF(I' . $rowNum . '="","",I' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AM' . $rowNum . '<>""),$AI$' . $rowNum . ':AM' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':H' . $rowNum . '<>""),$D$' . $rowNum . ':H' . $rowNum  . '))' . '+I' . $rowNum . '))))');
            /*VazaoBase07*/
            $sheet->setCellVAlue('AO' . $rowNum, '=IF(J' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AN' . $rowNum  . ')<=0,J' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AN' . $rowNum . '<>""),$AI$' . $rowNum . ':AN' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':I' . $rowNum . '<>""),$D$' . $rowNum . ':I' . $rowNum  . '))' . '+J' . $rowNum . '))>J' . $rowNum . ',IF(J' . $rowNum . '="","",J' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AN' . $rowNum . '<>""),$AI$' . $rowNum . ':AN' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':I' . $rowNum . '<>""),$D$' . $rowNum . ':I' . $rowNum  . '))' . '+J' . $rowNum . '))))');
            /*VazaoBase08*/
            $sheet->setCellVAlue('AP' . $rowNum, '=IF(K' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AO' . $rowNum  . ')<=0,K' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AO' . $rowNum . '<>""),$AI$' . $rowNum . ':AO' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':J' . $rowNum . '<>""),$D$' . $rowNum . ':J' . $rowNum  . '))' . '+K' . $rowNum . '))>K' . $rowNum . ',IF(K' . $rowNum . '="","",K' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AO' . $rowNum . '<>""),$AI$' . $rowNum . ':AO' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':J' . $rowNum . '<>""),$D$' . $rowNum . ':J' . $rowNum  . '))' . '+K' . $rowNum . '))))');
            /*VazaoBase09*/
            $sheet->setCellVAlue('AQ' . $rowNum, '=IF(L' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AP' . $rowNum  . ')<=0,L' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AP' . $rowNum . '<>""),$AI$' . $rowNum . ':AP' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':K' . $rowNum . '<>""),$D$' . $rowNum . ':K' . $rowNum  . '))' . '+L' . $rowNum . '))>L' . $rowNum . ',IF(L' . $rowNum . '="","",L' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AP' . $rowNum . '<>""),$AI$' . $rowNum . ':AP' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':K' . $rowNum . '<>""),$D$' . $rowNum . ':K' . $rowNum  . '))' . '+L' . $rowNum . '))))');
            /*VazaoBase10*/
            $sheet->setCellVAlue('AR' . $rowNum, '=IF(M' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AQ' . $rowNum  . ')<=0,M' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AQ' . $rowNum . '<>""),$AI$' . $rowNum . ':AQ' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':L' . $rowNum . '<>""),$D$' . $rowNum . ':L' . $rowNum  . '))' . '+M' . $rowNum . '))>M' . $rowNum . ',IF(M' . $rowNum . '="","",M' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AQ' . $rowNum . '<>""),$AI$' . $rowNum . ':AQ' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':L' . $rowNum . '<>""),$D$' . $rowNum . ':L' . $rowNum  . '))' . '+M' . $rowNum . '))))');
            /*VazaoBase11*/
            $sheet->setCellVAlue('AS' . $rowNum, '=IF(N' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AR' . $rowNum  . ')<=0,N' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AR' . $rowNum . '<>""),$AI$' . $rowNum . ':AR' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':M' . $rowNum . '<>""),$D$' . $rowNum . ':M' . $rowNum  . '))' . '+N' . $rowNum . '))>N' . $rowNum . ',IF(N' . $rowNum . '="","",N' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AR' . $rowNum . '<>""),$AI$' . $rowNum . ':AR' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':M' . $rowNum . '<>""),$D$' . $rowNum . ':M' . $rowNum  . '))' . '+N' . $rowNum . '))))');
            /*VazaoBase12*/
            $sheet->setCellVAlue('AT' . $rowNum, '=IF(O' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AS' . $rowNum  . ')<=0,O' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AS' . $rowNum . '<>""),$AI$' . $rowNum . ':AS' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':N' . $rowNum . '<>""),$D$' . $rowNum . ':N' . $rowNum  . '))' . '+O' . $rowNum . '))>O' . $rowNum . ',IF(O' . $rowNum . '="","",O' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AS' . $rowNum . '<>""),$AI$' . $rowNum . ':AS' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':N' . $rowNum . '<>""),$D$' . $rowNum . ':N' . $rowNum  . '))' . '+O' . $rowNum . '))))');
            /*VazaoBase13*/
            $sheet->setCellVAlue('AU' . $rowNum, '=IF(P' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AT' . $rowNum  . ')<=0,P' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AT' . $rowNum . '<>""),$AI$' . $rowNum . ':AT' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':O' . $rowNum . '<>""),$D$' . $rowNum . ':O' . $rowNum  . '))' . '+P' . $rowNum . '))>P' . $rowNum . ',IF(P' . $rowNum . '="","",P' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AT' . $rowNum . '<>""),$AI$' . $rowNum . ':AT' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':O' . $rowNum . '<>""),$D$' . $rowNum . ':O' . $rowNum  . '))' . '+P' . $rowNum . '))))');
            /*VazaoBase14*/
            $sheet->setCellVAlue('AV' . $rowNum, '=IF(Q' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AU' . $rowNum  . ')<=0,Q' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AU' . $rowNum . '<>""),$AI$' . $rowNum . ':AU' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':P' . $rowNum . '<>""),$D$' . $rowNum . ':P' . $rowNum  . '))' . '+Q' . $rowNum . '))>Q' . $rowNum . ',IF(Q' . $rowNum . '="","",Q' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AU' . $rowNum . '<>""),$AI$' . $rowNum . ':AU' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':P' . $rowNum . '<>""),$D$' . $rowNum . ':P' . $rowNum  . '))' . '+Q' . $rowNum . '))))');
            /*VazaoBase15*/
            $sheet->setCellVAlue('AW' . $rowNum, '=IF(R' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AV' . $rowNum  . ')<=0,R' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AV' . $rowNum . '<>""),$AI$' . $rowNum . ':AV' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':Q' . $rowNum . '<>""),$D$' . $rowNum . ':Q' . $rowNum  . '))' . '+R' . $rowNum . '))>R' . $rowNum . ',IF(R' . $rowNum . '="","",R' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AV' . $rowNum . '<>""),$AI$' . $rowNum . ':AV' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':Q' . $rowNum . '<>""),$D$' . $rowNum . ':Q' . $rowNum  . '))' . '+R' . $rowNum . '))))');
            /*VazaoBase16*/
            $sheet->setCellVAlue('AX' . $rowNum, '=IF(S' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AW' . $rowNum  . ')<=0,S' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AW' . $rowNum . '<>""),$AI$' . $rowNum . ':AW' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':R' . $rowNum . '<>""),$D$' . $rowNum . ':R' . $rowNum  . '))' . '+S' . $rowNum . '))>S' . $rowNum . ',IF(S' . $rowNum . '="","",S' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AW' . $rowNum . '<>""),$AI$' . $rowNum . ':AW' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':R' . $rowNum . '<>""),$D$' . $rowNum . ':R' . $rowNum  . '))' . '+S' . $rowNum . '))))');
            /*VazaoBase17*/
            $sheet->setCellVAlue('AY' . $rowNum, '=IF(T' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AX' . $rowNum  . ')<=0,T' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AX' . $rowNum . '<>""),$AI$' . $rowNum . ':AX' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':S' . $rowNum . '<>""),$D$' . $rowNum . ':S' . $rowNum  . '))' . '+T' . $rowNum . '))>T' . $rowNum . ',IF(T' . $rowNum . '="","",T' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AX' . $rowNum . '<>""),$AI$' . $rowNum . ':AX' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':S' . $rowNum . '<>""),$D$' . $rowNum . ':S' . $rowNum  . '))' . '+T' . $rowNum . '))))');
            /*VazaoBase18*/
            $sheet->setCellVAlue('AZ' . $rowNum, '=IF(U' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AY' . $rowNum  . ')<=0,U' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AY' . $rowNum . '<>""),$AI$' . $rowNum . ':AY' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':T' . $rowNum . '<>""),$D$' . $rowNum . ':T' . $rowNum  . '))' . '+U' . $rowNum . '))>U' . $rowNum . ',IF(U' . $rowNum . '="","",U' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AY' . $rowNum . '<>""),$AI$' . $rowNum . ':AY' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':T' . $rowNum . '<>""),$D$' . $rowNum . ':T' . $rowNum  . '))' . '+U' . $rowNum . '))))');
            /*VazaoBase19*/
            $sheet->setCellVAlue('BA' . $rowNum, '=IF(V' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AZ' . $rowNum  . ')<=0,V' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AZ' . $rowNum . '<>""),$AI$' . $rowNum . ':AZ' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':U' . $rowNum . '<>""),$D$' . $rowNum . ':U' . $rowNum  . '))' . '+V' . $rowNum . '))>V' . $rowNum . ',IF(V' . $rowNum . '="","",V' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':AZ' . $rowNum . '<>""),$AI$' . $rowNum . ':AZ' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':U' . $rowNum . '<>""),$D$' . $rowNum . ':U' . $rowNum  . '))' . '+V' . $rowNum . '))))');
            /*VazaoBase20*/
            $sheet->setCellVAlue('BB' . $rowNum, '=IF(W' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BA' . $rowNum  . ')<=0,W' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BA' . $rowNum . '<>""),$AI$' . $rowNum . ':BA' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':V' . $rowNum . '<>""),$D$' . $rowNum . ':V' . $rowNum  . '))' . '+W' . $rowNum . '))>W' . $rowNum . ',IF(W' . $rowNum . '="","",W' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BA' . $rowNum . '<>""),$AI$' . $rowNum . ':BA' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':V' . $rowNum . '<>""),$D$' . $rowNum . ':V' . $rowNum  . '))' . '+W' . $rowNum . '))))');
            /*VazaoBase21*/
            $sheet->setCellVAlue('BC' . $rowNum, '=IF(X' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BB' . $rowNum  . ')<=0,X' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BB' . $rowNum . '<>""),$AI$' . $rowNum . ':BB' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':W' . $rowNum . '<>""),$D$' . $rowNum . ':W' . $rowNum  . '))' . '+X' . $rowNum . '))>X' . $rowNum . ',IF(X' . $rowNum . '="","",X' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BB' . $rowNum . '<>""),$AI$' . $rowNum . ':BB' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':W' . $rowNum . '<>""),$D$' . $rowNum . ':W' . $rowNum  . '))' . '+X' . $rowNum . '))))');
            /*VazaoBase22*/
            $sheet->setCellVAlue('BD' . $rowNum, '=IF(Y' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BC' . $rowNum  . ')<=0,Y' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BC' . $rowNum . '<>""),$AI$' . $rowNum . ':BC' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':X' . $rowNum . '<>""),$D$' . $rowNum . ':X' . $rowNum  . '))' . '+Y' . $rowNum . '))>Y' . $rowNum . ',IF(Y' . $rowNum . '="","",Y' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BC' . $rowNum . '<>""),$AI$' . $rowNum . ':BC' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':X' . $rowNum . '<>""),$D$' . $rowNum . ':X' . $rowNum  . '))' . '+Y' . $rowNum . '))))');
            /*VazaoBase23*/
            $sheet->setCellVAlue('BE' . $rowNum, '=IF(Z' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BD' . $rowNum  . ')<=0,Z' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BD' . $rowNum . '<>""),$AI$' . $rowNum . ':BD' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':Y' . $rowNum . '<>""),$D$' . $rowNum . ':Y' . $rowNum  . '))' . '+Z' . $rowNum . '))>Z' . $rowNum . ',IF(Z' . $rowNum . '="","",Z' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BD' . $rowNum . '<>""),$AI$' . $rowNum . ':BD' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':Y' . $rowNum . '<>""),$D$' . $rowNum . ':Y' . $rowNum  . '))' . '+Z' . $rowNum . '))))');
            /*VazaoBase24*/
            $sheet->setCellVAlue('BF' . $rowNum, '=IF(AA' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BE' . $rowNum  . ')<=0,AA' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BE' . $rowNum . '<>""),$AI$' . $rowNum . ':BE' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':Z' . $rowNum . '<>""),$D$' . $rowNum . ':Z' . $rowNum  . '))' . '+AA' . $rowNum . '))>AA' . $rowNum . ',IF(AA' . $rowNum . '="","",AA' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BE' . $rowNum . '<>""),$AI$' . $rowNum . ':BE' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':Z' . $rowNum . '<>""),$D$' . $rowNum . ':Z' . $rowNum  . '))' . '+AA' . $rowNum . '))))');
            /*VazaoBase25*/
            $sheet->setCellVAlue('BG' . $rowNum, '=IF(AB' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BF' . $rowNum  . ')<=0,AB' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BF' . $rowNum . '<>""),$AI$' . $rowNum . ':BF' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AA' . $rowNum . '<>""),$D$' . $rowNum . ':AA' . $rowNum  . '))' . '+AB' . $rowNum . '))>AB' . $rowNum . ',IF(AB' . $rowNum . '="","",AB' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BF' . $rowNum . '<>""),$AI$' . $rowNum . ':BF' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AA' . $rowNum . '<>""),$D$' . $rowNum . ':AA' . $rowNum  . '))' . '+AB' . $rowNum . '))))');
            /*VazaoBase26*/
            $sheet->setCellVAlue('BH' . $rowNum, '=IF(AC' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BG' . $rowNum  . ')<=0,AC' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BG' . $rowNum . '<>""),$AI$' . $rowNum . ':BG' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AB' . $rowNum . '<>""),$D$' . $rowNum . ':AB' . $rowNum  . '))' . '+AC' . $rowNum . '))>AC' . $rowNum . ',IF(AC' . $rowNum . '="","",AC' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BG' . $rowNum . '<>""),$AI$' . $rowNum . ':BG' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AB' . $rowNum . '<>""),$D$' . $rowNum . ':AB' . $rowNum  . '))' . '+AC' . $rowNum . '))))');
            /*VazaoBase27*/
            $sheet->setCellVAlue('BI' . $rowNum, '=IF(AD' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BH' . $rowNum  . ')<=0,AD' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BH' . $rowNum . '<>""),$AI$' . $rowNum . ':BH' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AC' . $rowNum . '<>""),$D$' . $rowNum . ':AC' . $rowNum  . '))' . '+AD' . $rowNum . '))>AD' . $rowNum . ',IF(AD' . $rowNum . '="","",AD' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BH' . $rowNum . '<>""),$AI$' . $rowNum . ':BH' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AC' . $rowNum . '<>""),$D$' . $rowNum . ':AC' . $rowNum  . '))' . '+AD' . $rowNum . '))))');
            /*VazaoBase28*/
            $sheet->setCellVAlue('BJ' . $rowNum, '=IF(AE' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BI' . $rowNum  . ')<=0,AE' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BI' . $rowNum . '<>""),$AI$' . $rowNum . ':BI' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AD' . $rowNum . '<>""),$D$' . $rowNum . ':AD' . $rowNum  . '))' . '+AE' . $rowNum . '))>AE' . $rowNum . ',IF(AE' . $rowNum . '="","",AE' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BI' . $rowNum . '<>""),$AI$' . $rowNum . ':BI' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AD' . $rowNum . '<>""),$D$' . $rowNum . ':AD' . $rowNum  . '))' . '+AE' . $rowNum . '))))');
            /*VazaoBase29*/
            $sheet->setCellVAlue('BK' . $rowNum, '=IF(AF' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BJ' . $rowNum  . ')<=0,AF' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BJ' . $rowNum . '<>""),$AI$' . $rowNum . ':BJ' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AE' . $rowNum . '<>""),$D$' . $rowNum . ':AE' . $rowNum  . '))' . '+AF' . $rowNum . '))>AF' . $rowNum . ',IF(AF' . $rowNum . '="","",AF' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BJ' . $rowNum . '<>""),$AI$' . $rowNum . ':BJ' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AE' . $rowNum . '<>""),$D$' . $rowNum . ':AE' . $rowNum  . '))' . '+AF' . $rowNum . '))))');
            /*VazaoBase30*/
            $sheet->setCellVAlue('BL' . $rowNum, '=IF(AG' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BK' . $rowNum  . ')<=0,AG' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BK' . $rowNum . '<>""),$AI$' . $rowNum . ':BK' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AF' . $rowNum . '<>""),$D$' . $rowNum . ':AF' . $rowNum  . '))' . '+AG' . $rowNum . '))>AG' . $rowNum . ',IF(AG' . $rowNum . '="","",AG' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BK' . $rowNum . '<>""),$AI$' . $rowNum . ':BK' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AF' . $rowNum . '<>""),$D$' . $rowNum . ':AF' . $rowNum  . '))' . '+AG' . $rowNum . '))))');
            /*VazaoBase31*/
            $sheet->setCellVAlue('BM' . $rowNum, '=IF(AH' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BL' . $rowNum  . ')<=0,AH' . $rowNum . ',IF(((' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BL' . $rowNum . '<>""),$AI$' . $rowNum . ':BL' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AG' . $rowNum . '<>""),$D$' . $rowNum . ':AG' . $rowNum  . '))' . '+AH' . $rowNum . '))>AH' . $rowNum . ',IF(AH' . $rowNum . '="","",AH' . $rowNum . '),(' . $parA . '*(LOOKUP(1,1/($AI$' . $rowNum . ':BL' . $rowNum . '<>""),$AI$' . $rowNum . ':BL' . $rowNum  . ')))+((1-' . $parA . ')/2)*((LOOKUP(1,1/($D$' . $rowNum . ':AG' . $rowNum . '<>""),$D$' . $rowNum . ':AG' . $rowNum  . '))' . '+AH' . $rowNum . '))))');

            /*Recarga diária (mm/dia)*/
            /*RecargaDia01*/
            $sheet->setCellValue('BN' . $rowNum, '=IF(AI' . $rowNum . '="","",(AI' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia02*/
            $sheet->setCellValue('BO' . $rowNum, '=IF(AJ' . $rowNum . '="","",(AJ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia03*/
            $sheet->setCellValue('BP' . $rowNum, '=IF(AK' . $rowNum . '="","",(AK' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia04*/
            $sheet->setCellValue('BQ' . $rowNum, '=IF(AL' . $rowNum . '="","",(AL' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia05*/
            $sheet->setCellValue('BR' . $rowNum, '=IF(AM' . $rowNum . '="","",(AM' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia06*/
            $sheet->setCellValue('BS' . $rowNum, '=IF(AN' . $rowNum . '="","",(AN' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia07*/
            $sheet->setCellValue('BT' . $rowNum, '=IF(AO' . $rowNum . '="","",(AO' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia08*/
            $sheet->setCellValue('BU' . $rowNum, '=IF(AP' . $rowNum . '="","",(AP' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia09*/
            $sheet->setCellValue('BV' . $rowNum, '=IF(AQ' . $rowNum . '="","",(AQ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia10*/
            $sheet->setCellValue('BW' . $rowNum, '=IF(AR' . $rowNum . '="","",(AR' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia11*/
            $sheet->setCellValue('BX' . $rowNum, '=IF(AS' . $rowNum . '="","",(AS' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia12*/
            $sheet->setCellValue('BY' . $rowNum, '=IF(AT' . $rowNum . '="","",(AT' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia13*/
            $sheet->setCellValue('BZ' . $rowNum, '=IF(AU' . $rowNum . '="","",(AU' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia14*/
            $sheet->setCellValue('CA' . $rowNum, '=IF(AV' . $rowNum . '="","",(AV' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia15*/
            $sheet->setCellValue('CB' . $rowNum, '=IF(AW' . $rowNum . '="","",(AW' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia16*/
            $sheet->setCellValue('CC' . $rowNum, '=IF(AX' . $rowNum . '="","",(AX' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia17*/
            $sheet->setCellValue('CD' . $rowNum, '=IF(AY' . $rowNum . '="","",(AY' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia18*/
            $sheet->setCellValue('CE' . $rowNum, '=IF(AZ' . $rowNum . '="","",(AZ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia19*/
            $sheet->setCellValue('CF' . $rowNum, '=IF(BA' . $rowNum . '="","",(BA' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia20*/
            $sheet->setCellValue('CG' . $rowNum, '=IF(BB' . $rowNum . '="","",(BB' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia21*/
            $sheet->setCellValue('CH' . $rowNum, '=IF(BC' . $rowNum . '="","",(BC' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia22*/
            $sheet->setCellValue('CI' . $rowNum, '=IF(BD' . $rowNum . '="","",(BD' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia23*/
            $sheet->setCellValue('CJ' . $rowNum, '=IF(BE' . $rowNum . '="","",(BE' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia24*/
            $sheet->setCellValue('CK' . $rowNum, '=IF(BF' . $rowNum . '="","",(BF' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia25*/
            $sheet->setCellValue('CL' . $rowNum, '=IF(BG' . $rowNum . '="","",(BG' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia26*/
            $sheet->setCellValue('CM' . $rowNum, '=IF(BH' . $rowNum . '="","",(BH' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia27*/
            $sheet->setCellValue('CN' . $rowNum, '=IF(BI' . $rowNum . '="","",(BI' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia28*/
            $sheet->setCellValue('CO' . $rowNum, '=IF(BJ' . $rowNum . '="","",(BJ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia29*/
            $sheet->setCellValue('CP' . $rowNum, '=IF(BK' . $rowNum . '="","",(BK' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia30*/
            $sheet->setCellValue('CQ' . $rowNum, '=IF(BL' . $rowNum . '="","",(BL' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia30*/
            $sheet->setCellValue('CR' . $rowNum, '=IF(BM' . $rowNum . '="","",(BM' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');

            $n++;
            $nn++;
        }


        $writer = new Xlsx($spreadsheet);
        $writer->setPreCalculateFormulas(false);
        $codEstacao = ($_GET['codEstacao']);
        $fileName =  'FOCER-' . $codEstacao . '.Xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . urlencode($fileName) . '"');
        $writer->save('php://output');
    }

    /**
     * formata o array multidimensional associativo para um array multidimensional simples
     * onde o primeiro indice é um array com as keys e o restante são os dados
     * @param array $assoc_array
     * @return array
     */
    function format_array(array $assoc_array)
    {
        $array[] = array_keys($assoc_array[0]);
        for ($i = count($assoc_array) - 1; $i > 0; $i--) {
            $array[] = array_values($assoc_array[$i]);
        }
        return $array;
    }
    /**
     * método responsável por consultar os dados da ANA e exportar num arquivo xlsx
     */
    function export_data_from_ana()
    {
        $dataAsXmlString = get_xml_from_ana();
        if ($dataAsXmlString) {
            $dataAsAssocArray = xml_to_assoc_array($dataAsXmlString);
            $dataAsArray = format_array($dataAsAssocArray);
            export_to_xls($dataAsArray);
        }
    }
    export_data_from_ana();
} elseif ($radiobtn == 'metod3_ana') {

    /**
     * consulta os dados da ANA e retorna o conteúdo numa string formatado como xml 
     * @return string
     */
    function get_xml_from_ana()
    {
        $curl = curl_init();

        $codEstacao = ($_GET['codEstacao']);
        $dataInicio = ($_GET['dataInicio']);
        $dataFim = ($_GET['dataFim']);
        $nivelConsistencia = $_GET['nivelConsistencia'];
        $tipoDados = '3';

        curl_setopt_array($curl, [
            CURLOPT_URL => "http://telemetriaws1.ana.gov.br//ServiceANA.asmx/HidroSerieHistorica?codEstacao={$codEstacao}&dataInicio={$dataInicio}&dataFim={$dataFim}&tipoDados=3&nivelConsistencia={$nivelConsistencia}",
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_ENCODING => "",
            CURLOPT_MAXREDIRS => 10,
            CURLOPT_TIMEOUT => 30,
            CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
            CURLOPT_CUSTOMREQUEST => "GET",
            CURLOPT_POSTFIELDS => "",
        ]);

        $response = curl_exec($curl);
        $err = curl_error($curl);

        curl_close($curl);

        if ($err) {
            return false;
        }
        return (string) $response;
    }

    /**
     * Recebe o xml bruto, filtra apenas com os dados e retorna num array associativo
     * @param string $dataAsXmlString
     * @return array
     */
    function xml_to_assoc_array(string $dataAsXmlString)
    {
        $filters = [
            'EstacaoCodigo',
            'NivelConsistencia',
            'DataHora',
            'Vazao01',
            'Vazao02',
            'Vazao03',
            'Vazao04',
            'Vazao05',
            'Vazao06',
            'Vazao07',
            'Vazao08',
            'Vazao09',
            'Vazao10',
            'Vazao11',
            'Vazao12',
            'Vazao13',
            'Vazao14',
            'Vazao15',
            'Vazao16',
            'Vazao17',
            'Vazao18',
            'Vazao19',
            'Vazao20',
            'Vazao21',
            'Vazao22',
            'Vazao23',
            'Vazao24',
            'Vazao25',
            'Vazao26',
            'Vazao27',
            'Vazao28',
            'Vazao29',
            'Vazao30',
            'Vazao31',

        ];
        $regexFilter = implode('|', $filters);
        $array = [];

        $patern = '/\<SerieHistorica diffgr\:id\=\"SerieHistorica[0-9]+\" msdata\:rowOrder\=\"[0-9]+\"\>([^#]*?)\<\/SerieHistorica\>/';
        preg_match_all($patern, $dataAsXmlString, $filtered);
        $serieHistoricaList = $filtered[1];
        // $serieHistorica = $serieHistoricaList[0];
        foreach ($serieHistoricaList as $serieHistorica) {
            $patern = "/\<($regexFilter)?\>(.*?)\<\/.*\>/";
            preg_match_all($patern, $serieHistorica, $filtered);
            $titleList = $filtered[1];
            $valueList = $filtered[2];
            // var_dump($valueList);die;
            $serieHistoricaAsArray = [];
            for ($i = 0; $i < count($titleList); $i++) {
                $serieHistoricaAsArray[$titleList[$i]] = $valueList[$i];
            }
            $array[] = $serieHistoricaAsArray;
        }
        return $array;
    }

    /**
     * Recebe o array contendo os dados onde o primeiro indice é a header e o restante são os dados
     * @param array $data
     * @return bool
     */
    function export_to_xls(array $data)
    {

        $parA = ($_GET['formParametroA']);
        $parBFI = ($_GET['formParametroBFI']);
        $area = ($_GET['formAreaDrenagem']);

        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->fromArray($data);
        $sheet->setCellValue('AH1', 'Vazao31');
        $sheet->setCellValue('AI1', 'VazaoBase01');
        $sheet->setCellValue('AJ1', 'VazaoBase02');
        $sheet->setCellValue('AK1', 'VazaoBase03');
        $sheet->setCellValue('AL1', 'VazaoBase04');
        $sheet->setCellValue('AM1', 'VazaoBase05');
        $sheet->setCellValue('AN1', 'VazaoBase06');
        $sheet->setCellValue('AO1', 'VazaoBase07');
        $sheet->setCellValue('AP1', 'VazaoBase08');
        $sheet->setCellValue('AQ1', 'VazaoBase09');
        $sheet->setCellValue('AR1', 'VazaoBase10');
        $sheet->setCellValue('AS1', 'VazaoBase11');
        $sheet->setCellValue('AT1', 'VazaoBase12');
        $sheet->setCellValue('AU1', 'VazaoBase13');
        $sheet->setCellValue('AV1', 'VazaoBase14');
        $sheet->setCellValue('AW1', 'VazaoBase15');
        $sheet->setCellValue('AX1', 'VazaoBase16');
        $sheet->setCellValue('AY1', 'VazaoBase17');
        $sheet->setCellValue('AZ1', 'VazaoBase18');
        $sheet->setCellValue('BA1', 'VazaoBase19');
        $sheet->setCellValue('BB1', 'VazaoBase20');
        $sheet->setCellValue('BC1', 'VazaoBase21');
        $sheet->setCellValue('BD1', 'VazaoBase22');
        $sheet->setCellValue('BE1', 'VazaoBase23');
        $sheet->setCellValue('BF1', 'VazaoBase24');
        $sheet->setCellValue('BG1', 'VazaoBase25');
        $sheet->setCellValue('BH1', 'VazaoBase26');
        $sheet->setCellValue('BI1', 'VazaoBase27');
        $sheet->setCellValue('BJ1', 'VazaoBase28');
        $sheet->setCellValue('BK1', 'VazaoBase29');
        $sheet->setCellValue('BL1', 'VazaoBase30');
        $sheet->setCellValue('BM1', 'VazaoBase31');
        $sheet->setCellValue('BN1', 'RecargaDia01');
        $sheet->setCellValue('BO1', 'RecargaDia02');
        $sheet->setCellValue('BP1', 'RecargaDia03');
        $sheet->setCellValue('BQ1', 'RecargaDia04');
        $sheet->setCellValue('BR1', 'RecargaDia05');
        $sheet->setCellValue('BS1', 'RecargaDia06');
        $sheet->setCellValue('BT1', 'RecargaDia07');
        $sheet->setCellValue('BU1', 'RecargaDia08');
        $sheet->setCellValue('BV1', 'RecargaDia09');
        $sheet->setCellValue('BW1', 'RecargaDia10');
        $sheet->setCellValue('BX1', 'RecargaDia11');
        $sheet->setCellValue('BY1', 'RecargaDia12');
        $sheet->setCellValue('BZ1', 'RecargaDia13');
        $sheet->setCellValue('CA1', 'RecargaDia14');
        $sheet->setCellValue('CB1', 'RecargaDia15');
        $sheet->setCellValue('CC1', 'RecargaDia16');
        $sheet->setCellValue('CD1', 'RecargaDia17');
        $sheet->setCellValue('CE1', 'RecargaDia18');
        $sheet->setCellValue('CF1', 'RecargaDia19');
        $sheet->setCellValue('CG1', 'RecargaDia20');
        $sheet->setCellValue('CH1', 'RecargaDia21');
        $sheet->setCellValue('CI1', 'RecargaDia22');
        $sheet->setCellValue('CJ1', 'RecargaDia23');
        $sheet->setCellValue('CK1', 'RecargaDia24');
        $sheet->setCellValue('CL1', 'RecargaDia25');
        $sheet->setCellValue('CM1', 'RecargaDia26');
        $sheet->setCellValue('CN1', 'RecargaDia27');
        $sheet->setCellValue('CO1', 'RecargaDia28');
        $sheet->setCellValue('CP1', 'RecargaDia29');
        $sheet->setCellValue('CQ1', 'RecargaDia30');
        $sheet->setCellValue('CR1', 'RecargaDia31');

        $sheet->getStyle('D:CR')->getNumberFormat()->setFormatCode('0.00');

        $n = 1;
        $nn = 2;

        for ($i = 1; $i < count($data); $i++) {
            $rowNum = $n + 1;
            $rowNumMinus = $nn - 1;

            /*Vazão de base pelo método dos filtros númericos de Chapman e Maxwell (1996)*/
            /*VazaoBase01-Início da série histórica*/
            $sheet->setCellValue('AI2', '=IF(D2="","",D2)');
            /*VazaoBase01 - Após período inicial [DataInicio]*/
            $sheet->setCellValue('AI' . $rowNum, '=IF((AI' . $rowNumMinus . ':BM' . $rowNumMinus . ')="",IF(D' . $rowNum . '="","",D' . $rowNum . '),IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNumMinus . ':BM' . $rowNumMinus . '<>""),$AI$' . $rowNumMinus . ':BM' . $rowNumMinus  . '))+((1-' . $parA . ')/(2-' . $parA . '))*D' . $rowNum . ')>D' . $rowNum . ',IF(D' . $rowNum . '="","",D' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNumMinus . ':BM' . $rowNumMinus . '<>""),$AI$' . $rowNumMinus . ':BM' . $rowNumMinus  . '))+((1-' . $parA . ')/(2-' . $parA . '))*D' . $rowNum . '))');
            /*VazaoBase02*/
            $sheet->setCellValue('AJ' . $rowNum, '=IF(E' . $rowNum . '="","",IF(AI' . $rowNum . '="",E' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/(AI' . $rowNum . '<>""),AI' . $rowNum . '))+((1-' . $parA . ')/(2-' . $parA . '))*E' . $rowNum . ')>E' . $rowNum . ',IF(E' . $rowNum . '="","",E' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/(AI' . $rowNum . '<>""),AI' . $rowNum . '))+((1-' . $parA . ')/(2-' . $parA . '))*E' . $rowNum . ')))');
            /*VazaoBase03*/
            $sheet->setCellValue('AK' . $rowNum, '=IF(F' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AJ' . $rowNum  . ')<=0,F' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AJ' . $rowNum . '<>""),$AI$' . $rowNum . ':AJ' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*F' . $rowNum . ')>F' . $rowNum . ',IF(F' . $rowNum . '="","",F' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AJ' . $rowNum . '<>""),$AI$' . $rowNum . ':AJ' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*F' . $rowNum . ')))');
            /*VazaoBase04*/
            $sheet->setCellValue('AL' . $rowNum, '=IF(G' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AK' . $rowNum  . ')<=0,G' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AK' . $rowNum . '<>""),$AI$' . $rowNum . ':AK' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*G' . $rowNum . ')>G' . $rowNum . ',IF(G' . $rowNum . '="","",G' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AK' . $rowNum . '<>""),$AI$' . $rowNum . ':AK' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*G' . $rowNum . ')))');
            /*VazaoBase05*/
            $sheet->setCellValue('AM' . $rowNum, '=IF(H' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AL' . $rowNum  . ')<=0,H' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AL' . $rowNum . '<>""),$AI$' . $rowNum . ':AL' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*H' . $rowNum . ')>H' . $rowNum . ',IF(H' . $rowNum . '="","",H' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AL' . $rowNum . '<>""),$AI$' . $rowNum . ':AL' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*H' . $rowNum . ')))');
            /*VazaoBase06*/
            $sheet->setCellVAlue('AN' . $rowNum, '=IF(I' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AM' . $rowNum  . ')<=0,I' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AM' . $rowNum . '<>""),$AI$' . $rowNum . ':AM' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*I' . $rowNum . ')>I' . $rowNum . ',IF(I' . $rowNum . '="","",I' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AM' . $rowNum . '<>""),$AI$' . $rowNum . ':AM' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*I' . $rowNum . ')))');
            /*VazaoBase07*/
            $sheet->setCellVAlue('AO' . $rowNum, '=IF(J' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AN' . $rowNum  . ')<=0,J' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AN' . $rowNum . '<>""),$AI$' . $rowNum . ':AN' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*J' . $rowNum . ')>J' . $rowNum . ',IF(J' . $rowNum . '="","",J' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AN' . $rowNum . '<>""),$AI$' . $rowNum . ':AN' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*J' . $rowNum . ')))');
            /*VazaoBase08*/
            $sheet->setCellVAlue('AP' . $rowNum, '=IF(K' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AO' . $rowNum  . ')<=0,K' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AO' . $rowNum . '<>""),$AI$' . $rowNum . ':AO' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*K' . $rowNum . ')>K' . $rowNum . ',IF(K' . $rowNum . '="","",K' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AO' . $rowNum . '<>""),$AI$' . $rowNum . ':AO' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*K' . $rowNum . ')))');
            /*VazaoBase09*/
            $sheet->setCellVAlue('AQ' . $rowNum, '=IF(L' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AP' . $rowNum  . ')<=0,L' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AP' . $rowNum . '<>""),$AI$' . $rowNum . ':AP' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*L' . $rowNum . ')>L' . $rowNum . ',IF(L' . $rowNum . '="","",L' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AP' . $rowNum . '<>""),$AI$' . $rowNum . ':AP' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*L' . $rowNum . ')))');
            /*VazaoBase10*/
            $sheet->setCellVAlue('AR' . $rowNum, '=IF(M' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AQ' . $rowNum  . ')<=0,M' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AQ' . $rowNum . '<>""),$AI$' . $rowNum . ':AQ' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*M' . $rowNum . ')>M' . $rowNum . ',IF(M' . $rowNum . '="","",M' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AQ' . $rowNum . '<>""),$AI$' . $rowNum . ':AQ' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*M' . $rowNum . ')))');
            /*VazaoBase11*/
            $sheet->setCellVAlue('AS' . $rowNum, '=IF(N' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AR' . $rowNum  . ')<=0,N' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AR' . $rowNum . '<>""),$AI$' . $rowNum . ':AR' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*N' . $rowNum . ')>N' . $rowNum . ',IF(N' . $rowNum . '="","",N' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AR' . $rowNum . '<>""),$AI$' . $rowNum . ':AR' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*N' . $rowNum . ')))');
            /*VazaoBase12*/
            $sheet->setCellVAlue('AT' . $rowNum, '=IF(O' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AS' . $rowNum  . ')<=0,O' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AS' . $rowNum . '<>""),$AI$' . $rowNum . ':AS' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*O' . $rowNum . ')>O' . $rowNum . ',IF(O' . $rowNum . '="","",O' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AS' . $rowNum . '<>""),$AI$' . $rowNum . ':AS' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*O' . $rowNum . ')))');
            /*VazaoBase13*/
            $sheet->setCellVAlue('AU' . $rowNum, '=IF(P' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AT' . $rowNum  . ')<=0,P' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AT' . $rowNum . '<>""),$AI$' . $rowNum . ':AT' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*P' . $rowNum . ')>P' . $rowNum . ',IF(P' . $rowNum . '="","",P' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AT' . $rowNum . '<>""),$AI$' . $rowNum . ':AT' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*P' . $rowNum . ')))');
            /*VazaoBase14*/
            $sheet->setCellVAlue('AV' . $rowNum, '=IF(Q' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AU' . $rowNum  . ')<=0,Q' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AU' . $rowNum . '<>""),$AI$' . $rowNum . ':AU' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*Q' . $rowNum . ')>Q' . $rowNum . ',IF(Q' . $rowNum . '="","",Q' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AU' . $rowNum . '<>""),$AI$' . $rowNum . ':AU' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*Q' . $rowNum . ')))');
            /*VazaoBase15*/
            $sheet->setCellVAlue('AW' . $rowNum, '=IF(R' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AV' . $rowNum  . ')<=0,R' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AV' . $rowNum . '<>""),$AI$' . $rowNum . ':AV' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*R' . $rowNum . ')>R' . $rowNum . ',IF(R' . $rowNum . '="","",R' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AV' . $rowNum . '<>""),$AI$' . $rowNum . ':AV' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*R' . $rowNum . ')))');
            /*VazaoBase16*/
            $sheet->setCellVAlue('AX' . $rowNum, '=IF(S' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AW' . $rowNum  . ')<=0,S' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AW' . $rowNum . '<>""),$AI$' . $rowNum . ':AW' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*S' . $rowNum . ')>S' . $rowNum . ',IF(S' . $rowNum . '="","",S' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AW' . $rowNum . '<>""),$AI$' . $rowNum . ':AW' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*S' . $rowNum . ')))');
            /*VazaoBase17*/
            $sheet->setCellVAlue('AY' . $rowNum, '=IF(T' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AX' . $rowNum  . ')<=0,T' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AX' . $rowNum . '<>""),$AI$' . $rowNum . ':AX' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*T' . $rowNum . ')>T' . $rowNum . ',IF(T' . $rowNum . '="","",T' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AX' . $rowNum . '<>""),$AI$' . $rowNum . ':AX' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*T' . $rowNum . ')))');
            /*VazaoBase18*/
            $sheet->setCellVAlue('AZ' . $rowNum, '=IF(U' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AY' . $rowNum  . ')<=0,U' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AY' . $rowNum . '<>""),$AI$' . $rowNum . ':AY' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*U' . $rowNum . ')>U' . $rowNum . ',IF(U' . $rowNum . '="","",U' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AY' . $rowNum . '<>""),$AI$' . $rowNum . ':AY' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*U' . $rowNum . ')))');
            /*VazaoBase19*/
            $sheet->setCellVAlue('BA' . $rowNum, '=IF(V' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':AZ' . $rowNum  . ')<=0,V' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AZ' . $rowNum . '<>""),$AI$' . $rowNum . ':AZ' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*V' . $rowNum . ')>V' . $rowNum . ',IF(V' . $rowNum . '="","",V' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':AZ' . $rowNum . '<>""),$AI$' . $rowNum . ':AZ' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*V' . $rowNum . ')))');
            /*VazaoBase20*/
            $sheet->setCellVAlue('BB' . $rowNum, '=IF(W' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BA' . $rowNum  . ')<=0,W' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BA' . $rowNum . '<>""),$AI$' . $rowNum . ':BA' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*W' . $rowNum . ')>W' . $rowNum . ',IF(W' . $rowNum . '="","",W' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BA' . $rowNum . '<>""),$AI$' . $rowNum . ':BA' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*W' . $rowNum . ')))');
            /*VazaoBase21*/
            $sheet->setCellVAlue('BC' . $rowNum, '=IF(X' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BB' . $rowNum  . ')<=0,X' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BB' . $rowNum . '<>""),$AI$' . $rowNum . ':BB' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*X' . $rowNum . ')>X' . $rowNum . ',IF(X' . $rowNum . '="","",X' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BB' . $rowNum . '<>""),$AI$' . $rowNum . ':BB' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*X' . $rowNum . ')))');
            /*VazaoBase22*/
            $sheet->setCellVAlue('BD' . $rowNum, '=IF(Y' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BC' . $rowNum  . ')<=0,Y' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BC' . $rowNum . '<>""),$AI$' . $rowNum . ':BC' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*Y' . $rowNum . ')>Y' . $rowNum . ',IF(Y' . $rowNum . '="","",Y' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BC' . $rowNum . '<>""),$AI$' . $rowNum . ':BC' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*Y' . $rowNum . ')))');
            /*VazaoBase23*/
            $sheet->setCellVAlue('BE' . $rowNum, '=IF(Z' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BD' . $rowNum  . ')<=0,Z' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BD' . $rowNum . '<>""),$AI$' . $rowNum . ':BD' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*Z' . $rowNum . ')>Z' . $rowNum . ',IF(Z' . $rowNum . '="","",Z' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BD' . $rowNum . '<>""),$AI$' . $rowNum . ':BD' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*Z' . $rowNum . ')))');
            /*VazaoBase24*/
            $sheet->setCellVAlue('BF' . $rowNum, '=IF(AA' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BE' . $rowNum  . ')<=0,AA' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BE' . $rowNum . '<>""),$AI$' . $rowNum . ':BE' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AA' . $rowNum . ')>AA' . $rowNum . ',IF(AA' . $rowNum . '="","",AA' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BE' . $rowNum . '<>""),$AI$' . $rowNum . ':BE' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AA' . $rowNum . ')))');
            /*VazaoBase25*/
            $sheet->setCellVAlue('BG' . $rowNum, '=IF(AB' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BF' . $rowNum  . ')<=0,AB' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BF' . $rowNum . '<>""),$AI$' . $rowNum . ':BF' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AB' . $rowNum . ')>AB' . $rowNum . ',IF(AB' . $rowNum . '="","",AB' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BF' . $rowNum . '<>""),$AI$' . $rowNum . ':BF' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AB' . $rowNum . ')))');
            /*VazaoBase26*/
            $sheet->setCellVAlue('BH' . $rowNum, '=IF(AC' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BG' . $rowNum  . ')<=0,AC' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BG' . $rowNum . '<>""),$AI$' . $rowNum . ':BG' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AC' . $rowNum . ')>AC' . $rowNum . ',IF(AC' . $rowNum . '="","",AC' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BG' . $rowNum . '<>""),$AI$' . $rowNum . ':BG' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AC' . $rowNum . ')))');
            /*VazaoBase27*/
            $sheet->setCellVAlue('BI' . $rowNum, '=IF(AD' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BH' . $rowNum  . ')<=0,AD' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BH' . $rowNum . '<>""),$AI$' . $rowNum . ':BH' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AD' . $rowNum . ')>AD' . $rowNum . ',IF(AD' . $rowNum . '="","",AD' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BH' . $rowNum . '<>""),$AI$' . $rowNum . ':BH' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AD' . $rowNum . ')))');
            /*VazaoBase28*/
            $sheet->setCellVAlue('BJ' . $rowNum, '=IF(AE' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BI' . $rowNum  . ')<=0,AE' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BI' . $rowNum . '<>""),$AI$' . $rowNum . ':BI' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AE' . $rowNum . ')>AE' . $rowNum . ',IF(AE' . $rowNum . '="","",AE' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BI' . $rowNum . '<>""),$AI$' . $rowNum . ':BI' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AE' . $rowNum . ')))');
            /*VazaoBase29*/
            $sheet->setCellVAlue('BK' . $rowNum, '=IF(AF' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BJ' . $rowNum  . ')<=0,AF' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BJ' . $rowNum . '<>""),$AI$' . $rowNum . ':BJ' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AF' . $rowNum . ')>AF' . $rowNum . ',IF(AF' . $rowNum . '="","",AF' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BJ' . $rowNum . '<>""),$AI$' . $rowNum . ':BJ' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AF' . $rowNum . ')))');
            /*VazaoBase30*/
            $sheet->setCellVAlue('BL' . $rowNum, '=IF(AG' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BK' . $rowNum  . ')<=0,AG' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BK' . $rowNum . '<>""),$AI$' . $rowNum . ':BK' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AG' . $rowNum . ')>AG' . $rowNum . ',IF(AG' . $rowNum . '="","",AG' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BK' . $rowNum . '<>""),$AI$' . $rowNum . ':BK' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AG' . $rowNum . ')))');
            /*VazaoBase31*/
            $sheet->setCellVAlue('BM' . $rowNum, '=IF(AH' . $rowNum . '="","",IF(SUM($AI$' . $rowNum . ':BL' . $rowNum  . ')<=0,AH' . $rowNum . ',IF(((' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BL' . $rowNum . '<>""),$AI$' . $rowNum . ':BL' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AH' . $rowNum . ')>AH' . $rowNum . ',IF(AH' . $rowNum . '="","",AH' . $rowNum . '),(' . $parA . '/(2-' . $parA . '))*(LOOKUP(1,1/($AI$' . $rowNum . ':BL' . $rowNum . '<>""),$AI$' . $rowNum . ':BL' . $rowNum  . '))+((1-' . $parA . ')/(2-' . $parA . '))*AH' . $rowNum . ')))');

            /*Recarga diária (mm/dia)*/
            /*RecargaDia01*/
            $sheet->setCellValue('BN' . $rowNum, '=IF(AI' . $rowNum . '="","",(AI' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia02*/
            $sheet->setCellValue('BO' . $rowNum, '=IF(AJ' . $rowNum . '="","",(AJ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia03*/
            $sheet->setCellValue('BP' . $rowNum, '=IF(AK' . $rowNum . '="","",(AK' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia04*/
            $sheet->setCellValue('BQ' . $rowNum, '=IF(AL' . $rowNum . '="","",(AL' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia05*/
            $sheet->setCellValue('BR' . $rowNum, '=IF(AM' . $rowNum . '="","",(AM' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia06*/
            $sheet->setCellValue('BS' . $rowNum, '=IF(AN' . $rowNum . '="","",(AN' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia07*/
            $sheet->setCellValue('BT' . $rowNum, '=IF(AO' . $rowNum . '="","",(AO' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia08*/
            $sheet->setCellValue('BU' . $rowNum, '=IF(AP' . $rowNum . '="","",(AP' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia09*/
            $sheet->setCellValue('BV' . $rowNum, '=IF(AQ' . $rowNum . '="","",(AQ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia10*/
            $sheet->setCellValue('BW' . $rowNum, '=IF(AR' . $rowNum . '="","",(AR' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia11*/
            $sheet->setCellValue('BX' . $rowNum, '=IF(AS' . $rowNum . '="","",(AS' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia12*/
            $sheet->setCellValue('BY' . $rowNum, '=IF(AT' . $rowNum . '="","",(AT' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia13*/
            $sheet->setCellValue('BZ' . $rowNum, '=IF(AU' . $rowNum . '="","",(AU' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia14*/
            $sheet->setCellValue('CA' . $rowNum, '=IF(AV' . $rowNum . '="","",(AV' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia15*/
            $sheet->setCellValue('CB' . $rowNum, '=IF(AW' . $rowNum . '="","",(AW' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia16*/
            $sheet->setCellValue('CC' . $rowNum, '=IF(AX' . $rowNum . '="","",(AX' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia17*/
            $sheet->setCellValue('CD' . $rowNum, '=IF(AY' . $rowNum . '="","",(AY' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia18*/
            $sheet->setCellValue('CE' . $rowNum, '=IF(AZ' . $rowNum . '="","",(AZ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia19*/
            $sheet->setCellValue('CF' . $rowNum, '=IF(BA' . $rowNum . '="","",(BA' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia20*/
            $sheet->setCellValue('CG' . $rowNum, '=IF(BB' . $rowNum . '="","",(BB' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia21*/
            $sheet->setCellValue('CH' . $rowNum, '=IF(BC' . $rowNum . '="","",(BC' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia22*/
            $sheet->setCellValue('CI' . $rowNum, '=IF(BD' . $rowNum . '="","",(BD' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia23*/
            $sheet->setCellValue('CJ' . $rowNum, '=IF(BE' . $rowNum . '="","",(BE' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia24*/
            $sheet->setCellValue('CK' . $rowNum, '=IF(BF' . $rowNum . '="","",(BF' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia25*/
            $sheet->setCellValue('CL' . $rowNum, '=IF(BG' . $rowNum . '="","",(BG' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia26*/
            $sheet->setCellValue('CM' . $rowNum, '=IF(BH' . $rowNum . '="","",(BH' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia27*/
            $sheet->setCellValue('CN' . $rowNum, '=IF(BI' . $rowNum . '="","",(BI' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia28*/
            $sheet->setCellValue('CO' . $rowNum, '=IF(BJ' . $rowNum . '="","",(BJ' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia29*/
            $sheet->setCellValue('CP' . $rowNum, '=IF(BK' . $rowNum . '="","",(BK' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia30*/
            $sheet->setCellValue('CQ' . $rowNum, '=IF(BL' . $rowNum . '="","",(BL' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');
            /*RecargaDia30*/
            $sheet->setCellValue('CR' . $rowNum, '=IF(BM' . $rowNum . '="","",(BM' . $rowNum . '/(' . $area . '*1000000))*1000*(60*60*24))');

            $n++;
            $nn++;
        }


        $writer = new Xlsx($spreadsheet);
        $writer->setPreCalculateFormulas(false);
        $codEstacao = ($_GET['codEstacao']);
        $fileName =  'FOCER-' . $codEstacao . '.Xlsx';
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename="' . urlencode($fileName) . '"');
        $writer->save('php://output');
    }

    /**
     * formata o array multidimensional associativo para um array multidimensional simples
     * onde o primeiro indice é um array com as keys e o restante são os dados
     * @param array $assoc_array
     * @return array
     */
    function format_array(array $assoc_array)
    {
        $array[] = array_keys($assoc_array[0]);
        for ($i = count($assoc_array) - 1; $i > 0; $i--) {
            $array[] = array_values($assoc_array[$i]);
        }
        return $array;
    }
    /**
     * método responsável por consultar os dados da ANA e exportar num arquivo xlsx
     */
    function export_data_from_ana()
    {
        $dataAsXmlString = get_xml_from_ana();
        if ($dataAsXmlString) {
            $dataAsAssocArray = xml_to_assoc_array($dataAsXmlString);
            $dataAsArray = format_array($dataAsAssocArray);
            export_to_xls($dataAsArray);
        }
    }
    export_data_from_ana();
}
