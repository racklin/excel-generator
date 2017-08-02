<?php

namespace Racklin\ExcelGenerator;

use Illuminate\Support\ServiceProvider as LaravelServiceProvider;

/**
 * Class ExcelGeneratorServiceProvider
 *
 * @package Racklin\ExcelGenerator
 */
class ServiceProvider extends LaravelServiceProvider
{

    /**
     * Register the service provider.
     *
     * @return void
     */
    public function register()
    {
        $this->app->bind('excelgen', function () {
            return new ExcelGenerator();
        });
    }

    /**
     * Get the services provided by the provider.
     *
     * @return array
     */
    public function provides()
    {
        return [];
    }
}
