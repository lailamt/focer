<?php
define('dir_download', '../downloads/');
$arquivo = $_GET['arquivo'];
$arquivo = filter_var($arquivo, FILTER_SANITIZE_STRING);
$arquivo = basename($arquivo);
$caminho_download = dir_download . $arquivo;
if (!file_exists($caminho_download))
    die('Arquivo não existe!');
header('Content-type: octet/stream');
header('Content-disposition: attachment; filename="' . $arquivo . '";');
header('Content-Length: ' . filesize($caminho_download));
readfile($caminho_download);
exit;
?>