<?php

namespace golovchanskiy\parseTorg12;

use golovchanskiy\parseTorg12\models as models;
use golovchanskiy\parseTorg12\exceptions\ParseTorg12Exception;

/**
 * Разобрать товарную накладную по форме ТОРГ12 в формате Excel (.xls, .xlsx)
 *
 * @author Anton.Golovchanskiy <anton.golovchanskiy@gmail.com>
 */
class ParseTorg12
{

    /**
     * Путь к файлу
     *
     * @var string
     */
    private $filePath;

    /**
     * Допустимые значения ставки НДС
     * По умолчанию доступны: 0, 10, 18
     *
     * @var string
     */
    private $taxRateList;

    /**
     * Ставка НДС по-умолчанию (устаналивается, если не удалось определить ставку)
     * По умолчанию: 18
     *
     * @var string
     */
    private $defaultTaxRate;

    /**
     * Товарная накладная
     *
     * @var models\Invoice
     */
    public $invoice;

    /**
     * Атрибуты заголовка накладной
     *
     * @var array
     */
    private $settingsHeader = [
        'document_number' => [
            'label' => ['номер документа'],
            'shift_row' => 1,
        ],
        'document_date' => [
            'label' => ['дата составления'],
            'shift_row' => 1,
        ],
    ];

    /**
     * Синонимы для заголовков столбцов накладной
     *
     * @var array
     */
    private $settingsRow = [
        'num' => ['№', '№№', '№ п/п', 'номер по порядку'],
        'name' => ['название', 'наименование', 'наименование, характеристика, сорт, артикул товара'],
        'code' => ['код', 'isbn', 'ean', 'артикул', 'артикул поставщика', 'код товара поставщика', 'код (артикул)', 'штрих-код'],
        'cnt' => ['кол-во', 'количество', 'кол-во экз.', 'общее кол-во', 'количество (масса нетто)', 'коли-чество (масса нетто)'],
        'cnt_place' => ['мест, штук'],
        'not_cnt' => ['в одном месте'],
        'price_without_tax' => ['цена', 'цена без ндс', 'цена без ндс, руб.', 'цена без ндс руб.', 'цена без учета ндс', 'цена без учета ндс, руб.', 'цена без учета ндс руб.', 'цена, руб. коп.'],
        'price_with_tax' => ['цена с ндс, руб.', 'цена с ндс руб.', 'цена, руб.', 'цена руб.'],
        'sum_with_tax' => 'сумма.*с.*ндс', // regexp
        'tax_rate' => ['ндс, %', 'ндс %', 'ставка ндс, %', 'ставка ндс %', 'ставка ндс', 'ставка, %', 'ставка %'],
        'total' => ['всего по накладной'],
    ];

    /**
     * Активный лист документа
     *
     * @var \PHPExcel_Worksheet
     */
    private $worksheet;

    private $firstRow = 0; // номер строки, которую считаем началом заголовка
    private $startRow = NULL; // номер строки, которую считаем началом строк накладной
    private $highestRow; // номер последней строки
    private $highestColumn; // номер последнего столбца

    private $columnList = []; // координаты заголовков значащих столбцов
    private $rowsToProcess = []; // номера строк с нужными данными по накладной

    /**
     * @param string $filePath Путь к файлу
     * @param array $taxRateList Доступные ставки НДС
     * @param int $defaultTaxRate ставка НДС по умолчанию
     */
    public function __construct($filePath, array $taxRateList = [0, 10, 18], $defaultTaxRate = 18)
    {
        $this->filePath = $filePath;
        $this->taxRateList = $taxRateList;
        $this->defaultTaxRate = $defaultTaxRate;
    }

    /**
     * Разобрать накладную
     *
     * @throws ParseTorg12Exception
     */
    public function parse()
    {

        if (!file_exists($this->filePath)) {
            throw new ParseTorg12Exception('Указан некорректный путь к файлу накладной');
        }

        // читаем файл в формате Excel по форме ТОРГ12
        try {
            $objPHPExcel = \PHPExcel_IOFactory::load($this->filePath);
        } catch (\Exception $e) {
            $errorMsg = 'Невозможно прочитать загруженный файл: ' . $e->getMessage();
            throw new ParseTorg12Exception($errorMsg);
        }

        // создаем накладную
        $this->invoice = new models\Invoice();

        foreach ($objPHPExcel->getWorksheetIterator() as $worksheet) {
            $this->setWorksheet($worksheet);

            // очищаем список критических ошибок, т.к. накладная может быть не на первом листе
            $this->invoice->errors = [];

            // определяем последнюю строку документа
            $this->highestRow = $this->worksheet->getHighestRow();
            // определяем последний столбец документа
            $this->highestColumn = \PHPExcel_Cell::columnIndexFromString($this->worksheet->getHighestColumn());

            // разбираем заголовок накладной
            $this->parseHeader();

            // разбираем заголовок строк накладной
            $this->parseRowsHeader();

            // разбираем строки накладной, выкидываем дубли заголовка и т.п.
            $this->parseRows();

            // обрабатываем строки накладной
            $this->processRows();

            // если в накладной есть строки, то не обрабатываем остальные листы
            if (count($this->invoice->rows)) {

                // проверяем, что обработаны все строки накладной
                $lastRow = end($this->invoice->rows);
                if ($lastRow->num != count($this->invoice->rows)) {
                    $this->invoice->errors['count_rows'] = 'Порядковый номер последней строки накладной не совпадает с количеством обработанных строк';
                }

                break;
            }
        }

    }

    /**
     * Изменить активный лист
     *
     * @param \PHPExcel_Worksheet $worksheet
     */
    private function setWorksheet(\PHPExcel_Worksheet $worksheet)
    {
        $this->worksheet = $worksheet;
        $this->rowsToProcess = [];
        $this->columnList = [];
    }

    /**
     * Нормализуем содержимое ячейки
     *  - удаляем лишние пробелы
     *  - удаляем переносы строк
     *
     * @param string $cellValue Содержимое ячейки
     * @param bool $toLower Перевести все символы в нижний регистр
     * @return string
     */
    private function normalizeHeaderCellValue($cellValue, $toLower = true)
    {
        $cellValue = trim($cellValue);

        if ($toLower) {
            $cellValue = mb_strtolower($cellValue, 'UTF-8');
        }

        // удаляем странные пробелы, состоящие из 2 символов
        $cellValue = str_replace(chr(194) . chr(160), " ", $cellValue);
        // удаляем переносы строк
        $cellValue = str_replace("\n", " ", $cellValue);
        // удаляем ручные переносы строк "- "
        $cellValue = str_replace("- ", "", $cellValue);
        $cellValue = str_replace("-\n", "", $cellValue);
        $cellValue = str_replace("  ", " ", $cellValue);

        return $cellValue;
    }

    /**
     * Нормализуем содержимое ячейки
     *  - удаляем лишние пробелы
     *  - удаляем переносы строк
     *  - заменяем "," на ".", если в ячейке должно быть число
     *
     * @param string $cellValue Содержимое ячейки
     * @param bool $isNumber Число
     * @return string
     */
    private function normalizeCellValue($cellValue, $isNumber = false)
    {
        $cellValue = trim($cellValue);

        if ($isNumber) {
            $cellValue = str_replace(",", ".", $cellValue);
        }

        // удаляем странные пробелы, состоящие из 2 символов
        $cellValue = str_replace(chr(194) . chr(160), " ", $cellValue);
        // удаляем переносы строк
        $cellValue = str_replace("\n", " ", $cellValue);
        $cellValue = str_replace("  ", " ", $cellValue);

        return $cellValue;
    }

    /**
     * Получить атрибуты заголовка накладной
     *  - номер накладной
     *  - дата составления
     *
     * @return string
     * @throws ParseTorg12Exception
     */
    private function parseHeader()
    {
        $checkSell = function ($col, $row, $attribute) {
            $cellValue = $this->normalizeHeaderCellValue($this->worksheet->getCellByColumnAndRow($col, $row)->getValue());

            if (in_array($cellValue, $attribute['label'])) {
                // заголовок атрибута в одной ячейке
                $attributeValue = $this->normalizeCellValue($this->worksheet->getCellByColumnAndRow($col, $row + $attribute['shift_row'])->getValue());
                $this->firstRow = $row;
                return $attributeValue;
            } else {
                // заголовок атрибута разбит на две строки
                $nextValue = $this->normalizeHeaderCellValue($this->worksheet->getCellByColumnAndRow($col, $row + 1)->getValue());
                // считаем что два слова в заголовке всегда, если есть переносы - не распознается
                foreach ($attribute['label'] as $val) {
                    $multiRowHeader = explode(' ', $val);
                    if ($cellValue == $multiRowHeader[0] && $nextValue == $multiRowHeader[1]) {
                        $attributeValue = $this->normalizeCellValue($this->worksheet->getCellByColumnAndRow($col, $row + $attribute['shift_row'] + 1)->getValue());
                        $this->firstRow = $row;
                        return $attributeValue;
                    }
                }
            }

            return NULL;
        };

        // запоминаем координаты номера накладной
        for ($row = 0; $row <= $this->highestRow; $row++) {
            for ($col = 0; $col <= $this->highestColumn; $col++) {

                if (!empty($documentNumber) && !empty($documentDate)) {
                    break;
                }

                // номер
                if (empty($documentNumber)) {
                    $documentNumber = $checkSell($col, $row, $this->settingsHeader['document_number']);
                }

                // дата составления
                if (empty($documentDate)) {
                    $documentDate = $checkSell($col, $row, $this->settingsHeader['document_date']);
                }

            }
        }

        if ($documentNumber) {
            $this->invoice->number = $documentNumber;
        } else {
            $this->invoice->errors['invoice_number'] = 'Не найден номер накладной';
        }

        if (isset($documentDate)) {
            $documentTime = strtotime($documentDate);
            $this->invoice->date = date('Y-m-d', $documentTime);
        } else {
            $this->invoice->errors['invoice_date'] = 'Не найдена дата накладной';
        }

    }

    /**
     * Разобрать заголовок
     *
     * @throws ParseTorg12Exception
     */
    private function parseRowsHeader()
    {

        $match = function ($cellValue, $setting) {
            if (is_array($setting)) {
                return in_array($cellValue, $setting);
            } elseif (is_string($setting)) {
                return (bool)preg_match('#' . $setting . '#siu', $cellValue);
            } else {
                return false;
            }
        };

        /**
         * Запоминаем координаты первого заголовка
         *
         */
        for ($row = $this->firstRow; $row <= $this->highestRow; $row++) {
            for ($col = 0; $col <= $this->highestColumn; $col++) {

                $cellValue = $this->normalizeHeaderCellValue($this->worksheet->getCellByColumnAndRow($col, $row)->getValue());

                // нужна дополнительная проверка ячейки из следующей строки, т.к. заголовки дублируются
                $nextRowCellValue = $this->normalizeHeaderCellValue($this->worksheet->getCellByColumnAndRow($col, $row + 1)->getValue());

                if (!isset($this->columnList['num']) && $match($cellValue, $this->settingsRow['num'])) {

                    $this->columnList['num']['col'] = $col;
                    $this->columnList['num']['row'] = $row;

                } elseif (!isset($this->columnList['name']) && $match($cellValue, $this->settingsRow['name'])) {

                    $this->columnList['name']['col'] = $col;
                    $this->columnList['name']['row'] = $row;

                } elseif (!isset($this->columnList['code']) && $match($cellValue, $this->settingsRow['code'])) {

                    $this->columnList['code']['col'] = $col;
                    $this->columnList['code']['row'] = $row;

                } elseif (!isset($this->columnList['cnt']) && $match($cellValue, $this->settingsRow['cnt'])) {
                    // специальная обработка для количества, т.к. могут быть два одинаковых заголовка
                    if (!$match($nextRowCellValue, $this->settingsRow['not_cnt'])) {
                        $this->columnList['cnt']['col'] = $col;
                        $this->columnList['cnt']['row'] = $row;
                    }

                } elseif (!isset($this->columnList['cnt_place']) && $match($cellValue, $this->settingsRow['cnt_place'])) {

                    $this->columnList['cnt_place']['col'] = $col;
                    $this->columnList['cnt_place']['row'] = $row;

                } elseif (!isset($this->columnList['price_without_tax']) && $match($cellValue, $this->settingsRow['price_without_tax'])) {

                    $this->columnList['price_without_tax']['col'] = $col;
                    $this->columnList['price_without_tax']['row'] = $row;

                } elseif (!isset($this->columnList['price_with_tax']) && $match($cellValue, $this->settingsRow['price_with_tax'])) {

                    $this->columnList['price_with_tax']['col'] = $col;
                    $this->columnList['price_with_tax']['row'] = $row;

                } elseif (!isset($this->columnList['sum_with_tax']) && $match($cellValue, $this->settingsRow['sum_with_tax'])) {

                    $this->columnList['sum_with_tax']['col'] = $col;
                    $this->columnList['sum_with_tax']['row'] = $row;

                } elseif (!isset($this->columnList['tax_rate']) && $match($cellValue, $this->settingsRow['tax_rate'])) {

                    $this->columnList['tax_rate']['col'] = $col;
                    $this->columnList['tax_rate']['row'] = $row;

                }
            }
        }

        // проверяем корректность заголовка
        $this->checkRowsHeader();

        $this->startRow = $this->getMaxRowFromComplexHeader($this->columnList);
    }

    /**
     * Проверить корректность заголовка
     *
     * @throws ParseTorg12Exception
     */
    private function checkRowsHeader()
    {
        $headErrors = [];

        // проверяем наличие обязательных колонок
        if (empty($this->columnList)) {

            $headErrors[] = 'Необходимо указать названия столбцов';

        } elseif (!isset($this->columnList['num'])) {

            $msg = 'Необходимо добавить столбец, содержащий порядковый номер строки ("%s")';
            $headErrors[] = sprintf($msg, implode('"; "', $this->settingsRow['num']));

        } elseif (!isset($this->columnList['code'])) {

            $msg = 'Необходимо добавить столбец, содержащий код товара ("%s")';
            $headErrors[] = sprintf($msg, implode('"; "', $this->settingsRow['code']));

        } elseif (!isset($this->columnList['name'])) {

            $msg = 'Необходимо добавить столбец, содержащий название товара ("%s")';
            $headErrors[] = sprintf($msg, implode('"; "', $this->settingsRow['name']));

        } elseif (!isset($this->columnList['cnt']) && !isset($this->columnList['cnt_place'])) {

            $msg = 'Необходимо добавить столбец, содержащий количество товара ("%s")';
            $headErrors[] = sprintf($msg, implode('"; "', array_merge($this->settingsRow['cnt'], (array)$this->settingsRow['cnt_place'])));

        } elseif (!isset($this->columnList['price_without_tax'])) {

            $msg = 'Необходимо добавить столбец, содержащий цену товара без НДС ("%s")';
            $headErrors[] = sprintf($msg, implode('"; "', $this->settingsRow['price_without_tax']));

        } elseif (!isset($this->columnList['price_with_tax']) && !isset($this->columnList['sum_with_tax'])) {

            $msg = 'Необходимо добавить столбец, содержащий цену товара c НДС ("%s")';
            $headErrors[] = sprintf($msg, implode('"; "', array_merge($this->settingsRow['price_with_tax'], (array)$this->settingsRow['sum_with_tax'])));

        } elseif (!isset($this->columnList['tax_rate'])) {

            $msg = 'Необходимо добавить столбец, содержащий ставку НДС ("%s")';
            $headErrors[] = sprintf($msg, implode('"; "', $this->settingsRow['tax_rate']));

        }

        if ($headErrors) {
            throw new ParseTorg12Exception(implode("\n", $headErrors));
        }
    }

    /**
     * Разбираем накладную, определяем номера строк с позицими накладной
     *
     */
    private function parseRows()
    {
        $ws = $this->worksheet;

        for ($row = ($this->startRow + 1); $row <= $this->highestRow; $row++) {

            // прекращаем обработку, если попали в подвал накладной
            for ($col = 0; $col <= $this->highestColumn; $col++) {
                if (in_array($this->normalizeHeaderCellValue($ws->getCellByColumnAndRow($col, $row)->getValue()), $this->settingsRow['total'])) {
                    $this->highestRow = $row - 1;
                    return;
                }
            }

            $currentRow = [];
            $currentRow['num'] = $this->normalizeHeaderCellValue($ws->getCellByColumnAndRow($this->columnList['num']['col'], $row)->getValue());
            $currentRow['code'] = $this->normalizeHeaderCellValue($ws->getCellByColumnAndRow($this->columnList['code']['col'], $row)->getValue());

            // добавляем строку в обработку
            if ($this->validateRow($row, $currentRow)) {
                $this->rowsToProcess[] = $row;
            }
        }

    }

    /**
     * Проверить не является ли строка заголовком, т.к. ТОРГ12 может содержать несколько заголовков
     *
     * @param int $rowNumber Номер строки
     * @param array $currentRow Содержимое строки
     * @return bool
     */
    private function validateRow($rowNumber, $currentRow)
    {
        $row = [];
        $key = 1;

        for ($col = 0; $col <= $this->highestColumn; $col++) {
            $currentCell = $this->normalizeCellValue($this->worksheet->getCellByColumnAndRow($col, $rowNumber)->getValue());
            // запишем непустые значения в массив для текущей строки
            if ($currentCell) {
                $row[$key++] = $currentCell;
            }
        }

        // пропускаем строку с номерами столбцов
        if (
            count($row) > 2
            && ($row[1] == 1 && $row[2] == 2 && $row[3] == 3)
        ) {
            return false;
        }

        // пропускаем строку без порядкового номера
        if (!intval($currentRow['num'])) {
            return false;
        }

        // пропускаем повторные заголовки (достаточно, если в двух столбцах будет заголовок)
        if (
            in_array($currentRow['code'], $this->settingsRow['code']) ||
            in_array($currentRow['num'], $this->settingsRow['num'])
        ) {
            return false;
        }

        return true;
    }

    /**
     * Обработать валидные строки накладной
     *  - добавить строки в накладную
     *  - определить ошибки в строках накладной
     *
     */
    private function processRows()
    {
        $ws = $this->worksheet;

        for ($row = $this->startRow; $row <= $this->highestRow; ++$row) {

            // пропускаем строки, которые не надо обрабатывать
            if (!in_array($row, $this->rowsToProcess)) {
                continue;
            }

            $invoiceRow = new models\InvoiceRow();

            // порядковый номер
            $invoiceRow->num = (int)$this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['num']['col'], $row)->getValue());

            // код товара
            $invoiceRow->code = $this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['code']['col'], $row)->getValue());
            if (!$invoiceRow->code) {
                $invoiceRow->errors['code'] = 'Не указан код товара';
            }

            // название товара
            $invoiceRow->name = $this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['name']['col'], $row)->getValue());

            // количество
            if (isset($this->columnList['cnt'])) {
                $invoiceRow->cnt = (int)$this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['cnt']['col'], $row)->getValue(), true);
            }

            if (!$invoiceRow->cnt && isset($this->columnList['cnt_place'])) {
                $invoiceRow->cnt = (int)$this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['cnt_place']['col'], $row)->getValue(), true);
            }

            // цена без НДС
            $invoiceRow->price_without_tax = (float)$this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['price_without_tax']['col'], $row)->getValue(), true);
            if ($invoiceRow->price_without_tax) {
                $this->invoice->price_without_tax_sum += $invoiceRow->price_without_tax * $invoiceRow->cnt;
            }

            // цена c НДС
            if (isset($this->columnList['price_with_tax'])) {

                $invoiceRow->price_with_tax = (float)$this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['price_with_tax']['col'], $row)->getValue(), true);
                $this->invoice->price_with_tax_sum += $invoiceRow->price_with_tax * $invoiceRow->cnt;

            } elseif (isset($this->columnList['sum_with_tax'])) {

                $sumWithTax = $this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['sum_with_tax']['col'], $row)->getValue(), true);
                if ($sumWithTax) {
                    if ((int)$invoiceRow->cnt > 0) {
                        $invoiceRow->price_with_tax = round($sumWithTax / $invoiceRow->cnt, 4);
                    }
                    $this->invoice->price_with_tax_sum += $sumWithTax;
                }

            }

            if (!$invoiceRow->price_with_tax) {
                $invoiceRow->errors['price_with_tax'] = 'Не указана цена с учетом НДС';
            }

            // НДС
            $taxRate = $this->normalizeCellValue($ws->getCellByColumnAndRow($this->columnList['tax_rate']['col'], $row)->getValue(), true);
            $taxRate = str_replace('%', '', $taxRate);
            $taxRate = str_replace('без ндс', '0', strtolower($taxRate));
            $taxRate = intval($taxRate);
            if (in_array($taxRate, $this->taxRateList)) {
                $invoiceRow->tax_rate = $taxRate;
            } elseif (isset($this->defaultTaxRate)) {
                $invoiceRow->tax_rate = $this->defaultTaxRate;
                $invoiceRow->errors['tax_rate'] = sprintf('Установлено значение НДС по умолчанию: %d', $this->defaultTaxRate);
            } else {
                $invoiceRow->errors['tax_rate'] = sprintf('Значение НДС "%s" отсутсвует в списке доступных', $taxRate);
                $this->invoice->errors['tax_rate'] = 'В накладной присутсвует товар с некорректной ставкой НДС';
            }

            // проверка корректности указанной ставки НДС
            $calcPriceWithTax = round($invoiceRow->price_without_tax * (1 + $invoiceRow->tax_rate / 100), 2);
            $priceWithTax = round($invoiceRow->price_with_tax, 2);
            $diffPriceWithTax = abs($calcPriceWithTax - $priceWithTax);
            // погрешность 1 руб.
            if ($diffPriceWithTax > 1) {
                $invoiceRow->errors['diff_price_with_tax'] = sprintf('Некорректно указана ставка НДС (Цена с учётом НДС: %s, Рассчитанная цена с учетом НДС: %s', $priceWithTax, $calcPriceWithTax);
                $this->invoice->errors['diff_price_with_tax'] = 'В накладной присутсвует товар, по которому указана некорректная цена или ставка НДС';
            }

            // добавляем обработанную строку в накладную
            $this->invoice->rows[$invoiceRow->num] = $invoiceRow;
        }
    }

    /**
     * Получить номер последней строки многострочного заголовка
     *
     * @return int
     */
    private function getMaxRowFromComplexHeader()
    {
        $maxRow = 0;

        foreach ($this->columnList as $val) {
            $maxRow = ($val['row'] > $maxRow) ? $val['row'] : $maxRow;
        }

        // пропускаем строку с номерами столбцов
        $maxRow++;

        return $maxRow;
    }

} 