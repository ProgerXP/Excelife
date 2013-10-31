<?php
/*
  Excelife - parses Excel tables into trees using simple rules and chained calls.
  in public domain | by Proger_XP | https://github.com/ProgerXP/Excelife

  Supports PHP 5.3 and up. Requires excel_reader2 (included).
*/

if (is_file($config = __DIR__.DS.'bundle.php') and is_array($config = include $config)) {
  foreach ((array) array_get(reset($config), 'autoloads') as $type => $list) {
    if (is_array($list)) {
      foreach ($list as &$value) { $value = str_replace('(:bundle)', __DIR__, $value); }
      Autoloader::$type($list);
    }
  }
}
