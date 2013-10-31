# Excelife

**Excelife** parses Excel tables into trees using simple rules and chained calls.

It requires **excel_reader2** (included) and works out of the box with **PHP 5.3 and up** - simply include it and you're ready to go. Bundle for Laravel 3 is provided for easy autoloading.

**Unless you're using Laravel Excelife is just a single file (`excelife.php`) plus `excel_reader2.php`.**

## [Laravel bundle](http://bundles.laravel.com/bundle/excelife)
```
php artisan bundle:install excelife
```

## Example

The complete source code and sample data is available in `example` directory. You don't need it for Excelife to work. A snippet from there:

```
$rows = Excelife::make('data.xls')
  ->toAfter(0, '¹')
  ->match(function ($cells, $row) {
    if ($row->bgColor(1) === 'FFFF99') {
      return new Group($row, $cells[1]);
    }
  })
  ->match(function ($cells, $row) {
    if ($row->color(0) === 'FF9900') {
      return new Unavailable($row);
    } else {
      return new Product($row, array('sku', 'title', 'retail', 'wholesale'));
    }
  })
  ->parse()
  ->get();
```