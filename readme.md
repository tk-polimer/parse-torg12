ParseTorg12
=========

Разбор стандартной формы ТОРГ-12 в формате Excel (.xls, .xlsx).

Например, корректно разбирается накладная по форме ТОРГ-12 из 1С.

Установка
--------------

Для установки требуется [composer](https://getcomposer.org/).

Выполните команду

    php composer.phar require "golovchanskiy/parseTorg12" "*"

или добавьте в composer.json

    "require": {
        "golovchanskiy/parseTorg12": "*"
    },

Пример использования
--------------

    <?php
    require '../vendor/autoload.php';

    use \golovchanskiy\parseTorg12 as torg12;
    
    // указываем путь к файлу накладной по форме ТОРГ12
    $parseTorg12 = new torg12\ParseTorg12('./testTorg12.xls');

    try {
        // запускаем обработку накладной
        $parseTorg12->parse();
    } catch (torg12\ParseTorg12Exception $e) {
        // выводим ошибку обработки
        echo $e->getMessage();
    }

    if (!empty($parseTorg12->criticalErrors)) {
        // выводим ошибки в данных
        echo implode('<br>', $parseTorg12->criticalErrors);
    }

    // выводим результат работы
    var_dump((array)$parseTorg12->invoice);

Примеры накладных см. в папке example