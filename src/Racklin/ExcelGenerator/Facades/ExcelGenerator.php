<?php

namespace Racklin\ExcelGenerator\Facades;

use Illuminate\Support\Facades\Facade;

/**
 * Class ExcelGenerator
 *
 * @package Racklin\ExcelGenerator\Facades
 */
class ExcelGenerator extends Facade
{
    /**
     * @return string
     */
    protected static function getFacadeAccessor()
    {
        return 'excelgen';
    }
}
