# Simple Excel Generator for Laravel

Simple Excel Generator for Laravel using phpexcel library.

This package using json as template and you can pass php array as data to generate Excel xlsx. 
 
# Installation
```json
{
    "require": {
        "racklin/excel-generator": "dev-master"
    }
}
```

Next, add the service provider to `config/app.php`.

```php
'providers' => [
    //...
    Racklin\ExcelGenerator\ServiceProvider::class,
]

//...

'aliases' => [
	//...
	'ExcelGen' => Racklin\ExcelGenerator\Facades\ExcelGenerator::class
]

```

# Example
```
$excel = new ExcelGenerator();

$excel->generate('example_01.json', ["name"=>"rack", "cname"=>"阿土伯"], '/tmp/example.xlsx', 'F');
```
## Laravel Facade 
```
ExcelGen::generate('example_01.json', ["name"=>"rack", "cname"=>"阿土伯"], '/tmp/example.xlsx', 'F');
```

## Laravel version

Current package version works for Laravel 5+.

## License
MIT: https://racklin.mit-license.org/
