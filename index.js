'use strict';

var gutil = require('gulp-util'),
    through = require('through2'),
    XLSX = require('xlsx');

var debug = false;

var log = function(msg) {
    if (debug) {
        console.log(msg);
    }
};

/**
 * Generate headers from row.
 * @param worksheet {Object} - XLSX-Worksheet
 * @param headRow {number} - Header row index.
 * @param startColumn {string} - Column name, default 'A'
 * @return {Array} Return Array of Headers.
 * @example [{'A1': 'name'},{'B1': 'title'},{'C1': 'city'},...]
 */
var getHeaders = function(worksheet, headRow, startColumn) {
    var headers = [];

    for (var key in worksheet) {
        var cell = worksheet[key],
            match = /([A-Z]+)(\d+)/.exec(key);

        if (!match) {
            continue;
        }

        var row = +match[2]; // 1234
        var col = match[1]; // ABCD

        // If we've read past the header row, there is nothing else we're
        // going to do, so return the headers we've found.
        if (row > headRow) {
            break;
        }

        // If we aren't at the header row, then move to the next one.
        if (row < headRow) {
            continue;
        }

        // Pull in the data from the column, if we are in the right one.
        if (col >= startColumn) {
            var tmp = {};
            tmp[key] = cell.v;
            headers.push(tmp);
        }
    }

    return headers;
};

/**
 * Get value of headers from _key
 * @param _headers {Array} - Array of generated headers.
 * @param _headRow {number} - Index of header row. Default: 1.
 * @param _key {string} - Column name, for example 'A' (without index).
 * @return {*}
 */
var getHeadersValue = function(_headers, _headRow, _key) {
    var res = null;
    var key = _key + '' + _headRow;

    _headers.forEach(function(item) {
        if (item.hasOwnProperty(key)) {
            res = item[key];
        }
    });

    return res;
};

/**
 * Create dummy object from headers
 * @param {Array<Object>} headers
 */
var createItem = function(_headers) {
    var res = {},
        headers = JSON.parse(JSON.stringify(_headers));

    //create res object with all properties from headers
    headers.forEach(function(item, index) {
        var name = Object.keys(item)[0];
        res[item[name]] = '';
    });

    return JSON.parse(JSON.stringify(res));
};

/**
 * excel filename or workbook to json
 * @param fileName
 * @param headRow
 * @param startRow
 * @param startColumn
 * @returns {Array} json
 */
var toJson = function(fileName, headRow, startRow, startColumn) {
    var workbook;

    if (typeof fileName === 'string') {
        workbook = XLSX.readFile(fileName);
    } else {
        workbook = fileName;
    }

    var worksheet = workbook.Sheets[workbook.SheetNames[0]],
        json = [],
        namemap,
        curRow,
        value,
        row,
        col,
        cell,
        match,
        index,
        _c = 0,
        headers = getHeaders(worksheet, headRow, startColumn);

    namemap = createItem(headers);

    for (var key in worksheet) {
        if (worksheet.hasOwnProperty(key)) {
            cell = worksheet[key];
            match = /([A-Z]+)(\d+)/.exec(key);

            if (!match) {
                continue;
            }

            row = +match[2]; // 1234
            col = match[1]; // ABCD

            curRow = row - startRow;

            value = cell.v;

            if (col >= startColumn) {
                if (row < startRow) {
                    //continue;
                } else {
                    // check if we have a new line

                    if (_c < curRow) {
                        json[_c] = JSON.parse(JSON.stringify(namemap));
                        namemap = createItem(headers);
                        _c++;
                    }

                    namemap[getHeadersValue(headers, headRow, col)] = value;
                }
            }
        }
    }

    json[_c] = JSON.parse(JSON.stringify(namemap));

    return json;
};

module.exports = function(options) {
    options = options || {};
    debug = options.trace || false;

    return through.obj(function(file, enc, cb) {
        if (file.isNull()) {
            this.push(file);
            return cb();
        }

        if (file.isStream()) {
            this.emit('error', new gutil.PluginError(PLUGIN_NAME, 'Streaming not supported'));
            return cb();
        }

        var arr = [];
        for (var i = 0; i < file.contents.length; ++i) {
            arr[i] = String.fromCharCode(file.contents[i]);
        }

        var bString = arr.join("");

        /* Call XLSX */
        // TODO: add read-options for XLSX-module
        var workbook = XLSX.read(bString, {type: "binary"});
        file.contents = new Buffer(JSON.stringify(toJson(workbook, options.headRow || 1, options.valueRowStart || 2, options.startColumn || 'A')));

        log("Convert file: " + file.path);

        this.push(file);
        cb();
    });
};
