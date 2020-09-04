<?php

namespace golovchanskiy\parseTorg12\models;

/**
 * Строка товарной накладной
 */
class InvoiceRow extends Model
{

    /**
     * Порядковый номер
     *
     * @var int
     */
    public $num;

    /**
     * Код товара
     *
     * @var string
     */
    public $code;

    /**
     * Название товара
     *
     * @var string
     */
    public $name;

    /**
     * Ставка НДС
     *
     * @var int
     */
    public $tax_rate;

    /**
     * Цена с учетом НДС
     *
     * @var float
     */
    public $price_with_tax;

    /**
     * Цена без учета НДС
     *
     * @var float
     */
    public $price_without_tax;

    /**
     * Количество
     *
     * @var int
     */
    public $cnt;

    /**
     * Сумма с учетом НДС
     *
     * @var float
     */
    public $price_with_tax_sum;

}