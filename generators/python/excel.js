/**
 * @fileoverview Generating Python for math blocks.
 * @author q.neutron@gmail.com (Quynh Neutron)
 */
'use strict';

goog.provide('Blockly.Python.excel');

goog.require('Blockly.Python');


// If any new block imports any library, add that library name here.
Blockly.Python.addReservedWords('openpyxl');

Blockly.Python['get_workbook'] = function(block) {

  Blockly.Python.definitions_['import_openpyxl'] = 'import openpyxl';

  var filepath = Blockly.Python.valueToCode(block, 'FILEPATH', Blockly.Python.ORDER_ATOMIC);
  var code = 'openpyxl.load_workbook(' + filepath + ')';
  return [code, Blockly.Python.ORDER_ATOMIC];
};

Blockly.Python['get_sheet'] = function(block) {

  Blockly.Python.definitions_['import_openpyxl'] = 'import openpyxl';

  var excelbook = Blockly.Python.valueToCode(block, 'EXCELBOOK', Blockly.Python.ORDER_ATOMIC);
  var sheetname = Blockly.Python.valueToCode(block, 'SHEETNAME', Blockly.Python.ORDER_ATOMIC);
  var code = excelbook + '.get_sheet_by_name(' + sheetname + ')';
  return [code, Blockly.Python.ORDER_ATOMIC];
};

Blockly.Python['get_cell'] = function(block) {

  Blockly.Python.definitions_['import_openpyxl'] = 'import openpyxl';

  var sheet = Blockly.Python.valueToCode(block, 'EXCELSHEET', Blockly.Python.ORDER_ATOMIC);
  var cellname = Blockly.Python.valueToCode(block, 'CELLNAME', Blockly.Python.ORDER_ATOMIC);
  var code = sheet + '[' + cellname + ']';
  return [code, Blockly.Python.ORDER_ATOMIC];
};

Blockly.Python['get_cell_value'] = function(block) {

  Blockly.Python.definitions_['import_openpyxl'] = 'import openpyxl';

  var cell = Blockly.Python.valueToCode(block, 'EXCELCELL', Blockly.Python.ORDER_ATOMIC);
  var code = cell + '.value';
  return [code, Blockly.Python.ORDER_ATOMIC];
};

Blockly.Python['get_max_row'] = function(block) {

  Blockly.Python.definitions_['import_openpyxl'] = 'import openpyxl';

  var cell = Blockly.Python.valueToCode(block, 'EXCELSHEET', Blockly.Python.ORDER_ATOMIC);
  var code = cell + '.max_row';
  return [code, Blockly.Python.ORDER_ATOMIC];
};

Blockly.Python['get_max_column_number'] = function(block) {

  Blockly.Python.definitions_['import_openpyxl'] = 'import openpyxl';

  var cell = Blockly.Python.valueToCode(block, 'EXCELSHEET', Blockly.Python.ORDER_ATOMIC);
  var code = cell + '.max_column';
  return [code, Blockly.Python.ORDER_ATOMIC];
};

Blockly.Python['get_max_column_name'] = function(block) {

  Blockly.Python.definitions_['import_openpyxl'] = 'import openpyxl';

  var cell = Blockly.Python.valueToCode(block, 'EXCELSHEET', Blockly.Python.ORDER_ATOMIC);
  var code = 'get_column_letter(\'' + cell + '.max_column\')';
  return [code, Blockly.Python.ORDER_ATOMIC];
};
