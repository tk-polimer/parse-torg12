<?php

namespace golovchanskiy\parseTorg12;

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

}