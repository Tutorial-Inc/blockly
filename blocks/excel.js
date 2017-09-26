/*
 * @author Shiro Fukuda
*/


'use strict';

goog.provide('Blockly.Constants.Excels');

goog.require('Blockly.Blocks');

/**
 * Common HSV hue for all blocks in this category.
 * This should be the same as Blockly.Msg.COLOUR_HUE.
 * @readonly
 */
Blockly.Constants.Colour.HUE = 141;

Blockly.defineBlocksWithJsonArray([  // BEGIN JSON EXTRACT
  // Block for open excel book .
  {
    "type": "get_workbook",
    "message0": "%{BKY_EXCEL_GET_WORKBOOK}",
    "args0": [
      {
        "type": "input_value",
        "name": "FILEPATH",
        "input": "String"
      }
    ],
    "output": "ExcelBook",
    "colour": "%{BKY_EXCEL_HUE}",
  },
  {
    "type": "get_sheet",
    "message0": "%{BKY_EXCEL_GET_SHEET}",
    "args0": [
      {
        "type": "input_value",
        "name": "SHEETNAME",
        "input": "String"
      },
      {
        "type": "input_value",
        "name": "EXCELBOOK",
        "check": "ExcelBook"
      }
    ],
    "output": "ExcelSheet",
    "colour": "%{BKY_EXCEL_HUE}"
  },
  {
    "type": "get_cell",
    "message0": "%{BKY_EXCEL_GET_CELL}",
    "args0": [
      {
        "type": "input_value",
        "name": "CELLNAME",
        "input": "String"
      },
      {
        "type": "input_value",
        "name": "EXCELSHEET",
        "check": "ExcelSheet"
      }
    ],
    "output": "ExcelCell",
    "colour": "%{BKY_EXCEL_HUE}"
  },
  {
    "type": "get_cell_value",
    "message0": "%{BKY_EXCEL_GET_CELL_VALUE}",
    "args0": [{
      "type": "input_value",
      "name": "EXCELCELL",
      "input": "ExcelCell",
      "check": "ExcelCell"
    }],
    "output": "String",
    "colour": "%{BKY_EXCEL_HUE}"
  },
  {
    "type": "get_max_row",
    "message0": "%{BKY_EXCEL_GET_MAX_ROW}",
    "args0": [
      {
        "type": "input_value",
        "name": "EXCELSHEET",
        "input": "ExcelSheet",
        "check": "ExcelSheet"
      }
    ],
    "output": "Number",
    "colour": "%{BKY_EXCEL_HUE}"
  },
  {
    "type": "get_max_column_number",
    "message0": "%{BKY_EXCEL_GET_MAX_COLUMN_NUMBER}",
    "args0": [
      {
        "type": "input_value",
        "name": "EXCELSHEET",
        "input": "ExcelSheet",
        "check": "ExcelSheet"
      }
    ],
    "output": "Number",
    "colour": "%{BKY_EXCEL_HUE}"
  },
  {
    "type": "get_max_column_name",
    "message0": "%{BKY_EXCEL_GET_MAX_COLUMN_NAME}",
    "args0": [
      {
        "type": "input_value",
        "name": "EXCELSHEET",
        "input": "ExcelSheet",
        "check": "ExcelSheet"
      }
    ],
    "output": "String",
    "colour": "%{BKY_EXCEL_HUE}"
  }
]);
