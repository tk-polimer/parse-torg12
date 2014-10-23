<?php

namespace golovchanskiy\parseTorg12\models;


abstract class Model {

    /**
     * Список ошибок обработки
     *
     * @var array
     */
    public $errors = [];

    /**
     * Проверить валидность модели
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