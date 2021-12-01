"use strict";

var _interopRequireDefault = require("@babel/runtime/helpers/interopRequireDefault");

exports.__esModule = true;
exports.default = exports.build = exports.parseMetadata = exports.parse = void 0;

var _extends2 = _interopRequireDefault(
    require("@babel/runtime/helpers/extends")
);

var _xlsx = _interopRequireDefault(require("./xlsx"));

var _bufferFrom = _interopRequireDefault(require("buffer-from"));

var _helpers = require("./helpers");

var _workbook = _interopRequireDefault(require("./workbook"));

var _fs = require('fs');

var parse = function parse(mixed, options) {
    if (options === void 0) {
        options = {};
    }

    var workSheet = _xlsx.default[
        (0, _helpers.isString)(mixed) ? "readFile" : "read"
    ](mixed, options);

    return Object.keys(workSheet.Sheets).map(function (name) {
        var sheet = workSheet.Sheets[name];
        return {
            name,
            data: _xlsx.default.utils.sheet_to_json(sheet, {
                header: 1,
                raw: options.raw !== false,
                range: options.range ? options.range[name] : null,
            }),
        };
    });
};

exports.parse = parse;

var parseToHtml = function parse(mixed, options) {
    if (options === void 0) {
        options = {};
    }

    var workSheet = _xlsx.default[
        (0, _helpers.isString)(mixed) ? "readFile" : "read"
    ](mixed, options);

    return Object.keys(workSheet.Sheets).map(function (name) {
        var sheet = workSheet.Sheets[name];
        return {
            name,
            data: _xlsx.default.utils.sheet_to_html(sheet, {
                header: options.header,
                footer: options.footer,
            }),
        };
    });
};

exports.parseToHtml = parseToHtml;

var parseMetadata = function parseMetadata(mixed, options) {
    if (options === void 0) {
        options = {};
    }

    var workSheet = _xlsx.default[
        (0, _helpers.isString)(mixed) ? "readFile" : "read"
    ](mixed, options);

    return Object.keys(workSheet.Sheets).map(function (name) {
        var sheet = workSheet.Sheets[name];
        return {
            name,
            data: sheet["!ref"]
                ? _xlsx.default.utils.decode_range(sheet["!ref"])
                : null,
        };
    });
};

exports.parseMetadata = parseMetadata;

/*
    返回最终的表格数据，但是并不直接生成表格文件
*/
var build = function build(worksheets, options) {
    if (options === void 0) {
        options = {};
    }

    var defaults = {
        bookType: "xlsx",
        bookSST: false,
        type: "binary",
    };
    var workBook = new _workbook.default();
    worksheets.forEach(function (worksheet) {
        var sheetName = worksheet.name || "Sheet";
        var sheetOptions = worksheet.options || {};
        var sheetData = (0, _helpers.buildSheetFromMatrix)(
            worksheet.data || [],
            (0, _extends2.default)({}, options, sheetOptions)
        );
        workBook.SheetNames.push(sheetName);
        workBook.Sheets[sheetName] = sheetData;
    });

    var excelData = _xlsx.default.write(
        workBook,
        (0, _extends2.default)({}, defaults, options)
    );

    return excelData instanceof Buffer
        ? excelData
        : (0, _bufferFrom.default)(excelData, "binary");
};

exports.build = build;

/*
    直接生成最终的表格文件(同步)
*/
var write = function write(filename, worksheets, options) {
   var writeBuffer = build(worksheets, options);
   return _fs.writeFileSync(filename,writeBuffer)
};
exports.write = write;

/*
    直接生成最终的表格文件(异步)
*/
var writeAsync = function writeAsync(filename, worksheets, options,cb) {
   var writeBuffer = build(worksheets, options);
   return _fs.writeFile(filename, writeBuffer, cb);
};
exports.writeAsync = writeAsync;


var _XLSX = _xlsx.default;
exports._XLSX = _XLSX;

var _default = {
    parse,
    parseToHtml,
    parseMetadata,
    build,
    write,
    writeAsync,
    _XLSX,
};
exports.default = _default;