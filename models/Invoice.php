<?php

namespace golovchanskiy\parseTorg12\models;

/**
 * Товарная накладная
 */
class Invoice
{

    /**
     * Номер накладной
     *
     * @var string
     */
    public $number;

    /**
     * Дата накладной
     *
     * @var string
     */
    public $date;

    /**
     * Сумма накладной без учета НДС
     *
     * @var float
     */
    public $price_without_tax_sum = 0;

    /**
     * Сумма накладной с учетом НДС
     *
     * @var float
     */
    public $price_with_tax_sum = 0;

    /**
     * Строки накладной
     *
     * @var InvoiceRow[]
     */
    public $rows = [];

    /**
     * Список ошибок обработки накладной
     *
     * @var array
     */
    public $errors = [];

    /**
     * Проверить валидность накладной
     *
     * @return bool
     */
    public function isValid()
    {
        if (empty($this->errors)) {
            return true;
        } else {
            return false;
        }
    }

}