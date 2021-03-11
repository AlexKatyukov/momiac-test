<?php

require_once "vendor/autoload.php";

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

if(!empty($_FILES['file'])) {

    // если прилетел файл, то обрабатываем его

    $reader = new Xlsx();
    $table = $reader->load($_FILES['file']['tmp_name']);
    $sheet = $table->getActiveSheet();

    $letters = range('A', 'D');
    $numbers = range(1, 20);

    foreach ($letters as $l) {
        foreach ($numbers as $n) {
            if($sheet->getCell($l . $n)->getValue() % 2 != 0) {
                $sheet->setCellValue($l . $n, '');
            }
        }
    }

    header('Content-Type: application/vnd.ms-excel');
    header('Content-Disposition: attachment; filename="output.xlsx"');
    $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($table, "Xlsx");
    $writer->save("php://output");

} else {

    // если не прилетел файл, предлагаем загрузить

    echo <<<HTML
        <form action="" method="POST" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xls,.xlsx">
            <button type="submit">Загрузить файл</button>
        </form>
        <a href="https://github.com/AlexKatyukov/momiac-test" target="_blank">Смотреть на GitHub</a>
    HTML;

}