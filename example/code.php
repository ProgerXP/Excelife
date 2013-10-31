<?php
require_once '../excelife.php';

class Group extends ExcelifeGroup { }
class Product extends ExcelifeRow { }
class Unavailable extends ExcelifeRow { }

$rows = Excelife::make('data.xls')
  ->toAfter(0, '№')
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

foreach ($rows as $row) {
  echo get_class($row), ' ';

  if ($row instanceof Group) {
    echo "'{$row->title}'";
  } else {
    foreach ($row as $column => $value) {
      echo "$column: '$value' ";
    }
  }

  echo PHP_EOL;
}