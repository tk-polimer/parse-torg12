<html>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8"/>
    <title>parseTorg12: пример исполльзования</title>
</head>
<body>
<?php

require '../vendor/autoload.php';

ini_set('xdebug.var_display_max_depth', 5);
ini_set('xdebug.var_display_max_children', 256);
ini_set('xdebug.var_display_max_data', 1024);

use golovchanskiy\parseTorg12 as torg12;

$files = [
    './testTorg12.xls',
    './testTorg12.xlsx',
    './testTorg12_bad_path.xlsx',
    './testTorg12_big.xls',
];

foreach ($files as $filePath) {

    echo '<hr>';
    echo '<span style="color: darkblue; font-size: 20px; font-weight: bold;">' . $filePath . '</span><br>';

    // указываем путь к файлу накладной по форме ТОРГ12
    $parseTorg12 = new torg12\ParseTorg12($filePath);

    try {

        // запускаем обработку накладной
        $parseTorg12->parse();

        if (!$parseTorg12->invoice->isValid()) {
            echo '<pre>';
            echo '<span style="color: red; font-size: 20px; font-weight: bold;">При обработке накладной обнаружены ошибки:</span><br>';
            echo implode('<br>', $parseTorg12->invoice->errors);
            echo '</pre>';
            echo '<br>';
        }

        // выводим результат работы
        echo '<pre>';
        echo '<span style="color: green; font-size: 20px; font-weight: bold;">Результат:</span><br>';
        var_dump((array)$parseTorg12->invoice);
        echo '</pre>';

    } catch (torg12\exceptions\ParseTorg12Exception $e) {

        echo '<pre>';
        echo '<span style="color: red; font-size: 20px; font-weight: bold;">Ошибка:</span><br>';
        echo $e->getMessage();
        echo '</pre>';
        echo '<br>';

    }

}
?>
</body>
</html>