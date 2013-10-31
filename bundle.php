<?php
/*
  Excelife - parses Excel tables into trees using simple rules and chained calls.
  in public domain | by Proger_XP | https://github.com/ProgerXP/Excelife

  Supports PHP 5.3 and up. Requires excel_reader2 (included).
*/

return array(
  'excelife' => array(
    'autoloads' => array(
      'map' => array(
        'Spreadsheet_Excel_Reader'  => '(:bundle)/excel_reader2.php',
        'Excelife'                  => '(:bundle)/excelife.php',
        'ExcelifeRowInfo'           => '(:bundle)/excelife.php',
        'ExcelifeRow'               => '(:bundle)/excelife.php',
        'ExcelifeGroup'             => '(:bundle)/excelife.php',
      ),
    ),
  ),
);
