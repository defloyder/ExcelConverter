<?php
 

ini_set('display_errors', 1);
error_reporting(E_ALL);
session_start();
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpWord\PhpWord;
use ZipArchive;

ini_set('memory_limit', '2048M');
ini_set('max_execution_time', '600');

if (!isset($_SESSION['logs'])) {
    $_SESSION['logs'] = [];
}
if (!isset($_SESSION['progress'])) {
    $_SESSION['progress'] = 0;
}
if (!isset($_SESSION['start_time'])) {
    $_SESSION['start_time'] = time();
}

if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['file'])) {
    $_SESSION['logs'] = [];
    $_SESSION['progress'] = 0;
    $_SESSION['start_time'] = time();

    $uploadedFile = $_FILES['file'];
    $uploadedFilePath = 'uploads/' . basename($uploadedFile['name']);
    $outputDir = 'output/';
    $zipFileName = 'result.zip';
    $zipFilePath = $outputDir . $zipFileName;

    if (!is_dir('uploads/')) {
        mkdir('uploads/', 0777, true);
    }
    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0777, true);
    }

    array_map('unlink', glob("$outputDir/*.*"));

    if (move_uploaded_file($uploadedFile['tmp_name'], $uploadedFilePath)) {
        $reader = IOFactory::createReaderForFile($uploadedFilePath);
        $spreadsheet = $reader->load($uploadedFilePath);
        $sheet = $spreadsheet->getActiveSheet();
        $rows = $sheet->toArray();

        if (empty($rows) || count($rows) < 2) {
            echo json_encode(['status' => 'error', 'message' => 'Файл пуст или не содержит данных.']);
            exit;
        }

        $header = $rows[0];
        $data = array_slice($rows, 1);
        $groupedData = [];

        foreach ($data as $row) {
            $houseId = str_pad(preg_replace('/\D/', '', $row[0] ?? ''), 6, '0', STR_PAD_LEFT);
            if (!isset($groupedData[$houseId])) {
                $groupedData[$houseId] = [];
            }
            $groupedData[$houseId][] = $row;
        }

        $totalFiles = count($groupedData);
        $processed = 0;

        foreach ($groupedData as $houseId => $rows) {
            $processed++;
            $fileName = getFileNameFromAddress($header, $rows);
            saveBatchToFile($header, $rows, $outputDir . $fileName);
            $_SESSION['progress'] = round(($processed / $totalFiles) * 100);
        }

        $zip = new ZipArchive();
        if ($zip->open($zipFilePath, ZipArchive::CREATE | ZipArchive::OVERWRITE) === TRUE) {
            $files = scandir($outputDir);
            foreach ($files as $file) {
                if (pathinfo($file, PATHINFO_EXTENSION) === 'docx') {
                    $zip->addFile($outputDir . $file, basename($file));
                }
            }
            $zip->close();
        }

        if (file_exists($zipFilePath)) {
            echo json_encode(['status' => 'success', 'message' => 'Файл успешно обработан!', 'download_link' => $zipFileName]);
            exit;
        } else {
            echo json_encode(['status' => 'error', 'message' => 'Ошибка при создании архива']);
            exit;
        }
    } else {
        echo json_encode(['status' => 'error', 'message' => 'Ошибка загрузки файла']);
        exit;
    }
}

function getFileNameFromAddress($header, $rows) {
    $nameIndex = array_search('NAME', $header);
    $numIndex = array_search('NUM', $header);
    $korpIndex = array_search('KORP', $header);
    $strIndex = array_search('STR', $header);
    $cityIndex = array_search('CITY', $header);

    $name = $rows[0][$nameIndex] ?? '';
    $num = $rows[0][$numIndex] ?? '';
    $korp = $rows[0][$korpIndex] ?? '';
    $str = $rows[0][$strIndex] ?? '';
    $city = $rows[0][$cityIndex] ?? '';

    $addressParts = [];
    if ($city) {
        $addressParts[] = $city;
    }
    if ($name) {
        $addressParts[] = $name;
    }
    if ($num) {
        $addressParts[] = "д.$num";
    }
    if ($korp) {
        $addressParts[] = "корп.$korp";
    }
    if ($str) {
        $addressParts[] = "стр.$str";
    }

    $fileName = implode(' ', $addressParts);

    $fileName = preg_replace('/[<>:"\/\\|?*\x00-\x1F]/', '', $fileName); 
    $fileName = str_replace(' ', '', $fileName);
    $fileName = $fileName . '.docx';

    return $fileName;
}


function saveBatchToFile($header, $rows, $filePath) {
    $phpWord = new \PhpOffice\PhpWord\PhpWord();
    $section = $phpWord->addSection([
        'orientation' => 'landscape', 
    ]);

    $fontStyles = [
        'default' => ['name' => 'Times New Roman', 'size' => 12],
        'header' => ['name' => 'Times New Roman', 'size' => 14, 'bold' => true],
        'tableHeader' => ['name' => 'Times New Roman', 'size' => 12, 'bold' => true],
        'tableBody' => ['name' => 'Times New Roman', 'size' => 12],
        'address' => ['name' => 'Times New Roman', 'size' => 12, 'bold' => true],
    ];

    $paragraphStyles = [
        'default' => ['lineHeight' => 1.5],
        'noSpacing' => ['spaceAfter' => 0, 'lineHeight' => 1.5],
        'tableHeader' => ['lineHeight' => 1], 
        'tableBody' => ['lineHeight' => 1], 
        'leftAlign' => ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::LEFT], 
        'rightAlign' => ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::RIGHT], 
    ];

    $nameIndex = array_search('NAME', $header);
    $numIndex = array_search('NUM', $header);
    $korpIndex = array_search('KORP', $header);
    $strIndex = array_search('STR', $header);

    if ($nameIndex === false) {
        throw new Exception('Столбец "NAME" не найден в Excel файле.');
    }

    $nameValue = $rows[0][$nameIndex] ?? '';
    $numValue = $rows[0][$numIndex] ?? '';
    $korpValue = $rows[0][$korpIndex] ?? '';
    $strValue = $rows[0][$strIndex] ?? '';
    $cityValue = $rows[0]['CITY'] ?? '';

    $formattedAddress = [];
    if (!empty($nameValue)) {
        $nameValue = mb_convert_case($nameValue, MB_CASE_TITLE, "UTF-8");
        $formattedAddress[] = $nameValue;
    }
    if (!empty($numValue)) {
        $formattedAddress[] = "д. $numValue";
    }
    if (!empty($korpValue)) {
        $formattedAddress[] = "корп. $korpValue";
    }
    if (!empty($strValue)) {
        $formattedAddress[] = "стр. $strValue";
    }
    if (!empty($cityValue)) {
        $cityValue = mb_convert_case($cityValue, MB_CASE_TITLE, "UTF-8");
        $formattedAddress[] = $cityValue;
    }

    $headerText = implode(', ', $formattedAddress);

    $section->addText(
        "Описи сетей связи и оборудования АО «КОМКОР»",
        $fontStyles['header'],
        ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER]
    );
    $section->addTextBreak();

    $line = $section->addTextRun(); 
    $line->addText("г.Москва", $fontStyles['default']);
    $line->addText(str_repeat(' ', 200)); 
    $line->addText("«___»________", $fontStyles['default'], ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::RIGHT]);
    
    $section->addText(
        "Опись сети связи подготовлена оператором АО «КОМКОР» по объекту – многоквартирный дом, расположенный по адресу:",
        $fontStyles['default'],
        $paragraphStyles['noSpacing']
    );
    $section->addText($headerText, $fontStyles['address'], $paragraphStyles['noSpacing']);

    $removeColumns = [
        'NUM',
        'KORP',
        'STR',
        'NAME',
        'HOUSE_ID'
    ];

    $filteredHeader = array_diff($header, $removeColumns);
    array_unshift($filteredHeader, '№');

    $columnNamesMap = [
        'Элемент сети' => 'Элемент сети',
        'Перечень средств связи' => 'Перечень средств связи и линий связи',
        'Место монтажа сетей связи' => 'Место монтажа сетей связи',
        'Количество' => 'Количество энергопринимающих устройств',
        'Потребляемая мощьность(аВт)' => 'Потребляемая мощность одного энергопринимающего устройства/Максимальная потребляемая мощность (кВТ)',
        'Уровень напряжения(кВ)' => 'Уровень напряжения(кВ)',
        'Категория надежности' => 'Категория надежности энергоснабжения',
        'Перечень точек присоединения' => 'Перечень точек присоединения',
        'Информация о приборе учета' => 'Информация о приборе учета (при его наличии)',
    ];

    $table = $section->addTable(['borderSize' => 6, 'borderColor' => '000000', 'cellMargin' => 80]);

    $table->addRow();
    foreach ($filteredHeader as $headerText) {
        $displayName = $columnNamesMap[$headerText] ?? $headerText;
        $table->addCell(2000, ['valign' => 'top', 'align' => 'right'])->addText($displayName, $fontStyles['tableHeader'], $paragraphStyles['tableHeader']);
    }

    $rowNumber = 1;
    $totalElements = 0;
    $totalPower = 0.0;

    foreach ($rows as $row) {
        $table->addRow();
        $table->addCell(2000, ['valign' => 'center'])->addText($rowNumber, $fontStyles['tableBody'], $paragraphStyles['tableBody']);

        foreach ($filteredHeader as $columnName) {
            if ($columnName !== '№') {
                $index = array_search($columnName, $header);
                $value = $row[$index] ?? '-';

                if ($columnName === 'Количество') {
                    $totalElements += (int)$value;
                }

                if ($columnName === 'Потребляемая мощьность(аВт)') {
                    $totalPower += (float)$value / 1000;
                }

                $table->addCell(2000, ['valign' => 'center'])->addText($value, $fontStyles['tableBody'], $paragraphStyles['tableBody']);
            }
        }
        $rowNumber++;
    }

    $section->addText(
        "Общее количество элементов сети на доме составляет __{$totalElements}__ шт., общее количество потребляемой мощности составляет __" . number_format($totalPower, 3, '.', '') . "__ кВт.",
        $fontStyles['default'],
        $paragraphStyles['noSpacing']
    );
    $section->addText(
        "Опись подготовлена",
        $fontStyles['default'],
        $paragraphStyles['noSpacing']
    );

    $section->addText("Подпись _________________________________________", $fontStyles['default']);

    $phpWord->save($filePath, 'Word2007');
}




function logMessage($message) {
    $_SESSION['logs'][] = $message;
    error_log($message);
}
?>

<!DOCTYPE html>
<html lang="ru">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Загрузка Excel файла</title>
    <style>
        #progress {
            width: 100%;
            background-color: #f3f3f3;
            border: 1px solid #ccc;
            border-radius: 5px;
            overflow: hidden;
        }
        #progress-bar {
            height: 30px;
            width: 0;
            background-color: #4caf50;
            text-align: center;
            line-height: 30px;
            color: white;
        }
        #download-link {
            display: none;
            margin-top: 20px;
        }
    </style>
</head>
<body>
    <h1>Загрузите Excel файл</h1>
    <form id="upload-form" method="post" enctype="multipart/form-data">
        <input type="file" name="file" required>
        <button type="submit">Загрузить</button>
    </form>
    <div id="progress">
        <div id="progress-bar">0%</div>
    </div>
    <div id="logs"></div>
    <a id="download-link" href="#" download>Скачать архив</a>

    <script>
    const form = document.getElementById('upload-form');
    const progressBar = document.getElementById('progress-bar');
    const logsDiv = document.getElementById('logs');
    const downloadLink = document.getElementById('download-link');

    form.addEventListener('submit', function(event) {
        event.preventDefault();
        const formData = new FormData(form);
        const xhr = new XMLHttpRequest();
        xhr.open('POST', form .action, true);

        xhr.upload.addEventListener('progress', function(e) {
            if (e.lengthComputable) {
                const percentComplete = (e.loaded / e.total) * 100;
                progressBar.style.width = percentComplete + '%';
                progressBar.textContent = Math.round(percentComplete) + '%';
            }
        });

        xhr.onload = function() {
            if (xhr.status === 200) {
                const response = JSON.parse(xhr.responseText);
                console.log(response); ь
                if (response.status === 'error') {
                    logsDiv.innerHTML += '<p style="color:red;">' + response.message + '</p>';
                } else {
                    logsDiv.innerHTML += '<p style="color:green;">' + response.message + '</p>';
                    downloadLink.href = 'output/' + response.download_link;
                    downloadLink.style.display = 'block'; 
                }
            } else {
                logsDiv.innerHTML += '<p style="color:red;">Ошибка загрузки файла.</p>';
            }
        };

        xhr.send(formData);
    });
    </script>
</body>
</html>