<?php

require 'rec_dp.php';
//Relatório de recarga aquífera através de inserção de dados próprios
//========================================================
$RDP = new RecargaDP;
if (isset($_POST['submit']) and isset($_FILES['uploadFile']) and isset($_POST['metodo'])) {
    $radiobtn = $_POST['metodo'];
    if ($radiobtn == 'metod1_dp') {
        $RDP->Eckhardt();
    } elseif ($radiobtn == 'metod2_dp') {
        $RDP->LyneHollick();
    } elseif ($radiobtn == 'metod3_dp') {
        $RDP->ChapmanMaxwell();
    }
}
?>

<!DOCTYPE html>
<html lang="pt-br">

<head>
    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no" />
    <meta name="description" content="" />
    <meta name="author" content="" />
    <title>FOCER</title>
    <link href="css/styles.css" rel="stylesheet" />
    <link href="https://cdn.datatables.net/1.10.20/css/dataTables.bootstrap4.min.css" rel="stylesheet" crossorigin="anonymous" />
    <script src="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.1/js/all.min.js" crossorigin="anonymous"></script>

</head>

<body class="sb-nav-fixed">
    <nav class="sb-topnav navbar navbar-expand navbar-dark bg-dark">
        <a class="navbar-brand" href="index.html">FOCER v1.1.b</a>
        <button class="btn btn-link btn-sm order-1 order-lg-0" id="sidebarToggle" href="#"><i class="fas fa-bars"></i></button>

        <!-- Navbar-->
    </nav>
    <div id="layoutSidenav">
        <div id="layoutSidenav_nav">
            <nav class="sb-sidenav accordion sb-sidenav-dark" id="sidenavAccordion">
                <div class="sb-sidenav-menu">
                    <div class="nav">
                        <a class="nav-link collapsed" href="#" data-toggle="collapse" data-target="#collapseAbout" aria-expanded="false" aria-controls="collapseLayouts">
                            <div class="sb-nav-link-icon"><i class="far fa-window-maximize"></i></div>
                            SOBRE A FOCER
                            <div class="sb-sidenav-collapse-arrow"><i class="fas fa-angle-down"></i></div>
                        </a>
                        <div class="collapse" id="collapseAbout" aria-labelledby="headingOne" data-parent="#sidenavAccordion">
                            <nav class="sb-sidenav-menu-nested nav">
                                <a class="nav-link" href="index.html">Apresentação</a>
                                <a class="nav-link" href="profagua_ufba.html">ProfÁgua - Polo UFBA</a>
                            </nav>
                        </div>
                        <a class="nav-link collapsed" href="#" data-toggle="collapse" data-target="#collapseInterface" aria-expanded="false" aria-controls="collapseLayouts">
                            <div class="sb-nav-link-icon"><i class="fas fa-water"></i></div>
                            INTERFACE
                            <div class="sb-sidenav-collapse-arrow"><i class="fas fa-angle-down"></i></div>
                        </a>
                        <div class="collapse" id="collapseInterface" aria-labelledby="headingOne" data-parent="#sidenavAccordion">
                            <nav class="sb-sidenav-menu-nested nav">
                                <a class="nav-link" href="re_recarga.php">Relatório de recarga</a>
                                <a class="nav-link" href="mapa_interativo.html">Mapa interativo</a>
                            </nav>
                        </div>
                        <a class="nav-link" href="downloads.html">
                            <div class="sb-nav-link-icon"><i class="fas fa-download"></i></div>
                            DOWNLOADS
                        </a>
            </nav>
        </div>
        <div id="layoutSidenav_content">
            <main>
                <div class="card">
                    <div class="card-body">
                        <div class="card text-center">
                            <div class="card-header">
                                <ul class="nav nav-tabs card-header-tabs">
                                    <li class="nav-item">
                                        <a class="nav-link active" data-toggle="tab" href="#RecargaANA">Relatório de recarga (Base de dados da ANA)</a>
                                    </li>
                                    <li class="nav-item">
                                        <a class="nav-link" data-toggle="tab" href="#RecargaDados">Relatório de recarga (Utilizando dados prórpios)</a>
                                    </li>
                                </ul>
                            </div>
                            <div class="card-body">
                                <div class="tab-content">
                                    <div class="tab-pane fade show active" id="RecargaANA">
                                        <div class="card-body">
                                            <h5 class="card-title">Relatório de recarga aquífera</h5>
                                            <p class="card-text">A <b>FOCER</b> disponibiliza três tipos de filtros numéricos para separação do escoamento de base, permitindo assim estimar a taxa de recarga aquífera.
                                                Nesta seção você poderá estimar a taxa de recarga aquífera para as estações cadastradas na Rede Hidrometeorológica Nacional (RHN), escolhendo períodos históricos ou até mesmo a natureza dos dados
                                                (brutos ou consistidos). Disponibilizamos os parâmetros de entrada para aplicação dos filtros para 21 pontos de exutório, com referência às estações fluviométricas da RHN, essas informações
                                                podem ser consultadas através do mapa interativo. Para mais informações sobre as limitações de cada filtro e como são realizadas as estimativas de cada parâmetro de entrada, visite o manual
                                                do usuário na seção de metadados.
                                            </p>
                                            <form action="rec_ana.php" method="$_GET">
                                                <br>
                                                <div class="row justify-content-center">
                                                    <div class="col-auto">
                                                        <center>
                                                            <p><b>Parâmetros (ANA)</b></p>
                                                        </center>
                                                        <div class="form-inline">
                                                            <div class="col-auto">
                                                                <label for="formAreaDrenagem">Código da estação</label>
                                                                <input type="text" class="form-control" name="codEstacao" id="codEstacao" required />
                                                            </div>
                                                            <div class="col-auto">
                                                                <label for="formParametroA">Data (início)</label>
                                                                <input type="date" class="form-control" name="dataInicio" id="dataInicio" required />
                                                            </div>
                                                            <div class="col-auto">
                                                                <label for="formParametroBFI">Data (fim)</label>
                                                                <input type="date" class="form-control" name="dataFim" id="dataFim" required />
                                                            </div>
                                                            <div class="col-auto">
                                                                <label for="nivelConsistencia">Tipo de dados</label>
                                                                <select class="custom-select" id="tipoDeDados" name="nivelConsistencia" id="nivelConsistencia" required>
                                                                    <option value="1">Dados brutos </option>
                                                                    <option value="2">Dados consistidos </option>
                                                                </select>
                                                            </div>
                                                        </div>
                                                    </div>
                                                    <div class="col-auto">
                                                        <center>
                                                            <br>
                                                            <p><b>Métodos (Filtros numéricos)</b></p>
                                                        </center>
                                                        <div class="form-check form-check-inline">
                                                            <input type="radio" class="form-check-input" name="metodo" value="metod1_ana" id="metod1_ana" required />
                                                            <label class="form-check-label" for="metod1_ana">
                                                                Eckhardt (2005)
                                                            </label>
                                                            <input type="radio" class="form-check-input" name="metodo" value="metod2_ana" id="metod2_ana" required />
                                                            <label class="form-check-label" for="metod2_ana">
                                                                Lyne & Hollick (1979)
                                                            </label>
                                                            <input type="radio" class="form-check-input" name="metodo" value="metod3_ana" id="metod3_ana" required />
                                                            <label class="form-check-label" for="metod3_ana">
                                                                Chapman & Maxwell (1996)
                                                            </label>
                                                        </div>
                                                    </div>
                                                    <div class="col-auto">
                                                        <center>
                                                            <br>
                                                            <p><b>Parâmetros (Filtros numéricos)</b></p>
                                                        </center>
                                                        <div class="form-inline">
                                                            <div class="col-sm">
                                                                <label for="formAreaDrenagem">Área de drenagem (km²)</label>
                                                                <input type="number" class="form-control" name="formAreaDrenagem" id="formAreaDrenagem" min="0" step="any" required />
                                                            </div>
                                                            <div class="col-sm">
                                                                <label for="formParametroA">Constante de recessão (α)</label>
                                                                <input type="number" class="form-control" name="formParametroA" id="formParametroA" min="0" max="1" step="any" />
                                                            </div>
                                                            <div class="col-sm">
                                                                <label for="formParametroBFI">Base Flown Index (BFI)</label>
                                                                <input type="number" class="form-control" name="formParametroBFI" id="formParametroBFI" min="0" max="1" step="any" />
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                                <br>
                                                <div class="form-group">
                                                    <center><button class="btn btn-primary" type="submit" name="submit2" id="submit2"> Gerar relatório</button></center>
                                                </div>
                                            </form>
                                        </div>
                                    </div>
                                    <div class="tab-pane fade" id="RecargaDados">
                                        <div class="card-body">
                                            <h5 class="card-title">Relatório de recarga aquífera</h5>
                                            <p class="card-text">A <b>FOCER</b> disponibiliza três tipos de filtros numéricos para separação do escoamento de base, permitindo assim estimar a taxa de recarga aquífera.
                                                Nesta seção você poderá estimar a taxa de recarga aquífera a partir de dados próprios. Realize o upload do arquivo com os dados organizados em duas colunas, a primeira coluna
                                                com os dados de tempo, em dias, e a segunda coluna com os dados de vazão, em m³/s. O arquivo deve ser enviado em formato .csv e os dados devem ser organizados de forma
                                                contínua sem falhas. Para mais informações sobre as limitações de cada filtro, como os dados são devem ser organizados e como são realizadas as estimativas de cada parâmetro 
                                                de entrada, visite o manual do usuário na seção de metadados.
                                            </p>
                                            <form action="#" method="POST" enctype="multipart/form-data">
                                                <br>
                                                <div class="row justify-content-center">
                                                    <div class="col-auto">
                                                        <center>
                                                            <p><b>Métodos (Filtros numéricos)</b></p>
                                                        </center>
                                                        <div class="form-check form-check-inline">
                                                            <input type="radio" class="form-check-input" name="metodo" value="metod1_dp" id="metod1_dp" required />
                                                            <label class="form-check-label" for="metod1_dp">
                                                                Eckhardt (2005)
                                                            </label>
                                                            <input type="radio" class="form-check-input" name="metodo" value="metod2_dp" id="metod2_dp" required />
                                                            <label class="form-check-label" for="metod2_dp">
                                                                Lyne & Hollick (1979)
                                                            </label>
                                                            <input type="radio" class="form-check-input" name="metodo" value="metod3_dp" id="metod3_dp" required />
                                                            <label class="form-check-label" for="metod3_dp">
                                                                Chapman & Maxwell (1996)
                                                            </label>
                                                        </div>
                                                    </div>
                                                    <div class="col-auto">
                                                        <center>
                                                            <br>
                                                            <p><b>Parâmetros (Filtros numéricos)</b></p>
                                                        </center>
                                                        <div class="form-inline">
                                                            <div class="col-sm">
                                                                <label for="formAreaDrenagem">Área de drenagem (km²)</label>
                                                                <input type="number" class="form-control" name="formAreaDrenagem" id="formAreaDrenagem" min="0" step="any" />
                                                            </div>
                                                            <div class="col-sm">
                                                                <label for="formParametroA">Constante de recessão (α)</label>
                                                                <input type="number" class="form-control" name="formParametroA" id="formParametroA" min="0" max="1" step="any" />
                                                            </div>
                                                            <div class="col-sm">
                                                                <label for="formParametroBFI">Base Flown Index (BFI)</label>
                                                                <input type="number" class="form-control" name="formParametroBFI" id="formParametroBFI" min="0" max="1" step="any" />
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
                                                <br>
                                                <div class="row justify-content-center">
                                                    <div class="col-auto">
                                                        <label class="DragAndDrop" for="uploadFile">
                                                            <div class="d-flex justify-content-center"><i class="fas fa-file-upload fa-2x" style="display:none"></i></div>
                                                            <input type="file" name="uploadFile" class="fileInput" id="uploadFile" accept=".csv" required></input>
                                                        </label>
                                                        <div class="form-group">
                                                            <center><button class="btn btn-primary" type="submit" name="submit"> Gerar relatório</button></center>
                                                        </div>
                                                    </div>
                                            </form>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </main>
            <footer class="py-4 bg-light mt-auto">
                <div class="container-fluid">
                    <div class="d-flex align-items-center justify-content-between small">
                        <div class="text-muted">Este obra está licenciada com uma Licença Creative Commons de atribuição
                            não comercial 4.0 internacional.</div>
                        <div>
                            <a href="#">Política de privacidade</a>
                            &middot;
                            <a href="#">Termos &amp; Condições</a>
                        </div>
                    </div>
                </div>
            </footer>
        </div>
    </div>
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js" crossorigin="anonymous"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.3/dist/js/bootstrap.bundle.min.js" crossorigin="anonymous"></script>
    <script src="js/scripts.js"></script>
   


</body>

</html>