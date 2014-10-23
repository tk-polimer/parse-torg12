ParseTorg12
=========

Разбор стандартной формы ТОРГ-12 в формате Excel (.xls, .xlsx).

Например, корректно разбирается накладная по форме ТОРГ-12 из 1С.

Установка
--------------

Для установки требуется [composer](https://getcomposer.org/).

Выполните команду

    php composer.phar require "golovchanskiy/parse-torg12" "dev-master"

или добавьте в composer.json

    "require": {
        "golovchanskiy/parse-torg12": "dev-master"
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
        
        if (!$parseTorg12->invoice->isValid()) {
            // выводим ошибки обработки накладной
            echo implode('<br>', $parseTorg12->invoice->errors);
        }
    
        // выводим результат работы
        var_dump((array)$parseTorg12->invoice);
            
    } catch (torg12\exceptions\ParseTorg12Exception $e) {
    
        // выводим ошибку обработки
        echo $e->getMessage();
        
    }

Примеры накладных см. в папке example

Результат
--------------

В реузльтате получаем следующие данные:

### Накладная:
* Номер
* Дата составления

### Строка накладной (товар):
* Порядковый номер
* Код товара
* Название товара
* Ставка НДС
* Цена с учетом НДС
* Цена без учета НДС
* Количество
