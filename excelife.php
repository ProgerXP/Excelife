<?php
/*
  Excelife - parses Excel tables into trees using simple rules and chained calls.
  in public domain | by Proger_XP | https://github.com/ProgerXP/Excelife

  Supports PHP 5.3 and up. Requires excel_reader2 (included).
*/

class Excelife implements IteratorAggregate, Countable {
  //= Spreadsheet_Excel_Reader
  protected $xls;

  //= int cached sheet dimensions
  protected $rowCount, $colCount;

  //= array of ExcelifeRowInfo
  protected $rows = array();

  //= array of object
  protected $parsed = array();

  // function (array $cells, ExcelifeRowInfo $row, Excelife $this)
  // Returns null or an object (usually an ExcelifeRow child).
  //
  //= array of array of (callable, (bool) is_object)
  protected $matchers = array();

  //= int current row number (0-based)
  protected $current = 0;
  //= null infinite, int last row number to parse
  protected $end;

  static function wrongArg($func, $msg) {
    throw new InvalidArgumentException(__CLASS__.".$func(): $msg");
  }

  static function fail($msg, $class = 'Exception') {
    throw new $class(__CLASS__.": $msg");
  }

  static function loadReader() {
    if (!class_exists('Spreadsheet_Excel_Reader')) {
      $file = __DIR__.'/excel_reader2.php';
      if (is_file($file)) { include_once $file; }
    }
  }

  //= Spreadsheet_Excel_Reader
  static function readerFrom($xls) {
    static::loadReader();

    if ($xls instanceof Spreadsheet_Excel_Reader) {
      return $xls;
    } elseif (!is_scalar($xls)) {
      static::wrongArg(__FUNCTION__, '$xls must be .xls file path or reader instance.');
    } elseif (!is_file($xls) or !is_readable($xls)) {
      static::wrongArg(__FUNCTION__, "file [$xls] doesn't exist or isn't readable.");
    } else {
      $errors = static::readerErrors(false);
      $xls = new Spreadsheet_Excel_Reader($xls);
      static::readerErrors($errors);
      return $xls;
    }
  }

  // Spreadsheet_Excel_Reader productes lots of E_NOTICE's while working.
  //
  //? $oldLevel = readerErrors(false)
  //    // remember current settings and do what's necessary
  //  readerErrors($oldLevel)
  //    // restore them and continue normal execution
  static function readerErrors($enable) {
    if ($enable) {
      error_reporting($enable);
    } else {
      if (($current = error_reporting()) & E_NOTICE) {
        error_reporting($current & ~E_NOTICE);
      }

      return $current;
    }
  }

  static function make($xls) {
    return new static($xls);
  }

  //* $xls str file name, Spreadsheet_Excel_Reader
  function __construct($xls) {
    $this->xls = static::readerFrom($xls);
    $this->refresh();
  }

  function refresh() {
    $this->rowCount = $this->xls->rowCount();
    $this->colCount = $this->xls->colCount();
    $this->rows = array();

    return $this;
  }

  //= Spreadsheet_Excel_Reader
  function reader() {
    return $this->xls;
  }

  //= int
  function rowCount() {
    return $this->rowCount;
  }

  //= int
  function colCount() {
    return $this->colCount;
  }

  //= str text at given coord (0-based indexes)
  function at($x, $y, $trim = true) {
    $value = $this->xls->val($y + 1, $x + 1);
    return $trim ? trim($value) : $value;
  }

  // function ()
  // Clears currently set matchers.
  //
  // function (callable $callback)
  // Adds a new matched. A matcher is executed once per each row until end() and
  // should return either an object or a falsy value (null, false, 0, '').
  function match($callback) {
    if ($callback === null) {
      $this->matchers = array();
    } else {
      $this->matchers[] = array($callback, is_object($callback));
    }

    return $this;
  }

  // function ($y)
  // parse() will start processing from row $y (0-based).
  //
  // function ($x, $caption)
  // Find a row having $caption in its $x'th column (0-based) and set the starting
  // point there.
  function to($num, $caption = null) {
    if (func_num_args() < 2) {
      $y = $num;
    } else {
      $y = -1;
      $end = $this->rowCount;

      while ($this->at($num, ++$y) !== $caption) {
        if ($y >= $end) {
          ++$num;
          static::fail("to() cannot find a row with [$caption] at column $num.");
        }
      }
    }

    $this->current = $y;
    return $this;
  }

  function toAfter($x, $caption) {
    return $this->to($x, $caption)->skip();
  }

  function skip($count = 1) {
    $this->current += $count;
    return $this;
  }

  function back($count = 1) {
    $this->current -= $count;
    return $this;
  }

  function current() {
    return $this->current;
  }

  // function ()
  //= int last row to be parse()'d, null process until rowCount()
  //
  // function ($end)
  // Set or unset the ending row.
  //* $end int set new value, null set to rowCount()
  function end($y = null) {
    func_num_args() and $this->end = $y;
    return func_num_args() ? $this : $this->end;
  }

  // Clears this object of all previous parsing effects.
  function reset() {
    $this->parsed = array();
    return $this;
  }

  // Matches specified rows (see to(), end()) using assigned matchers. After it
  // returns get() produced results.
  function parse() {
    $current = &$this->current;

    $end = $this->end;
    $end = $end ? min($end, $this->rowCount - 1) : ($this->rowCount - 1);

    for (; $current <= $end; ++$current) {
      $obj = $this->matchRow($current);
      $obj and $this->parsed[] = $obj;
    }

    return $this;
  }

  //= null, object
  protected function matchRow($y) {
    $row = $this->row($y);
    $cells = $row->cells();

    foreach ($this->matchers as $matcher) {
      // up to 20X faster calling of Closure's - useful on long Excel files.
      $obj = $matcher[1]
        ? $matcher[0]($cells, $row, $this)
        : call_user_func($matcher[0], $cells, $row, $this);

      if (is_object($obj)) {
        return $obj;
      } elseif ($obj) {
        static::fail('matcher should return either an object or a falsy value,'.
                     ' received '.gettype($obj).'.', 'UnexpectedValueException');
      }
    }
  }

  // Returns the same object for the same row (0-based $y) until refresh() is called.
  //= ExcelifeRowInfo
  function row($y) {
    $row = &$this->rows[$y];
    $row or $row = new ExcelifeRowInfo($this->xls, $y);
    return $row;
  }

  //= array of mixed parsed objects
  function get() {
    return $this->parsed;
  }

  function getIterator() {
    return new ArrayIterator($this->parsed);
  }

  function count() {
    return count($this->parsed);
  }
}

// Provides access to Excel's style information, cell values, etc. Can be used
// independently of Excelife object.
class ExcelifeRowInfo {
  public $xls;            //= Spreadsheet_Excel_Reader
  public $y;              //= int 0-based

  //* $y int - 0-based.
  function __construct(Spreadsheet_Excel_Reader $xls, $y) {
    $this->xls = $xls;
    $this->y = $y;
  }

  // Virtual methods that can be called (see Excel Reader's docs):
  //
  //   type (number, date, unknown), raw, hyperlink (URL), rowspan, colspan,
  //   format (string), align (right, center, '' (left)), bgColor (RRGGBB),
  //   color (RRGGBB), bold, italic, underline, height (pixels), font (name),
  //   style (CSS), borderLeft/Right/Top/Bottom, borderLeftColor/etc. (RRGGBB)
  //
  // Note: colors are returned without leading hash (#) and in upper case.
  //
  //= scalar
  function __call($method, $params) {
    $params += array(0);
    $result = $this->xls->$method($this->y + 1, $params[0] + 1);

    if ($method === 'bgColor') {
      $result = $this->xls->rawColor($result);
    }

    if ($method[strlen($method) - 1] === 'r' and
        !strcasecmp(substr($method, -5), 'color')) {
      return strtoupper( ltrim($result, '#') );
    } else {
      return $result;
    }
  }

  //= array of scalar trimmed cell values at this row
  function cells($fromX = 0, $toX = 999) {
    $y = $this->y + 1;
    ++$fromX;
    $toX = min($toX, $this->xls->colCount() - 1) + 1;

    $cells = array();

    for (; $fromX <= $toX; ++$fromX) {
      $cells[] = trim( $this->xls->val($y, $fromX) );
    }

    return $cells;
  }

  // Retrieves single cell's value.
  function at($x, $trim = true) {
    $value = $this->xls->val($this->y + 1, $x + 1);
    return $trim ? trim($value) : $value;
  }
}

// Base object returned by a matcher. It's not necessary to use it but it provides
// quick convenient way of mapping row's values onto named values and normalizing them.
class ExcelifeRow implements IteratorAggregate, Countable {
  // List of column names that should be always present - they're normalized
  // (from null) even if they are not present in initial columns.
  //= array of str
  public $names = array();

  public $row;            //= ExcelifeRowInfo
  public $columns;        //= hash of str 'name' => 'cell value'

  static function make(ExcelifeRowInfo $row, $columns = null) {
    return new static($row, $columns);
  }

  //* $row - parent row this object belongs to; used to retrieve cell values.
  //* $columns array - cell names to assign properties using; see fill().
  function __construct(ExcelifeRowInfo $row, $columns = null) {
    $this->row = $row;
    $this->fill($columns);
  }

  // Maps the contents of the row this object is connected to onto column names.
  //* $columns array of str - if there are more names than columns they receive
  //  'null' values; if there are less - the rest is ignored. Empty names ('')
  //  are ignored as well.
  function fill($columns) {
    $columns = (array) $columns;
    $cells = $this->row->cells();

    while (count($columns) > count($cells)) {
      $columns[] = null;
    }

    $cells = array_slice($cells, 0, count($columns));
    $this->columns = $columns ? array_combine($columns, $cells) : array();
    unset( $this->columns[''] );

    return $this;
  }

  // function ( [null] )
  // Normalize all columns.
  //
  // function (str $name)
  // Normalize column with given name and return its new value.
  //
  // function (str $name, mixed $value)
  // Return  normalized $value as the value of column $name. New value isn't saved.
  function normalize($name = null, $value = null) {
    if ($name === null) {
      $names = array_unique( array_merge(array_keys($this->columns), $this->names) );
      array_map(array($this, 'normalize'), $names);
      return $this;
    } else {
      func_num_args() == 1 and $value = &$this->columns[$name];

      $func = 'normalize_'.strtolower($name);
      method_exists($this, $func) and $value = $this->$func($value);

      return $value;
    }
  }

  // function ($name, $value)
  // Set column $name to $value.
  //
  // function (hash $columns)
  // Sets column values from the array.
  function set($name, $value = null) {
    is_array($name) or $name = array($name => $value);
    foreach ($name as $name => $value) { $this->columns[$name] = $value; }
    return $this;
  }

  function __get($name) {
    return @$this->columns[$name];
  }

  function __isset($name) {
    return isset( $this->columns[$name] );
  }

  function __set($name, $value) {
    $this->columns[$name] = $value;
  }

  function __unset($name) {
    unset( $this->columns[$name] );
  }

  function getIterator() {
    return new ArrayIterator($this->columns);
  }

  function count() {
    return count($this->columns);
  }
}

// Usually used to separate groups of rows from each other. Is identical to normal
// ExcelifeRow but is meant to have a single property - just $title.
class ExcelifeGroup extends ExcelifeRow {
  public $title;          //= str trim()'ed if assigned in the constructor

  //* $row - parent row this object belongs to; used to retrieve cell values.
  //* $title array column names, str group title
  function __construct(ExcelifeRowInfo $row, $title = null) {
    $names = is_array($title) ? $title : null;
    parent::__construct($row, $names);
    $names or $this->title = $title;
  }
}
