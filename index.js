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
    var headers = [],
        tmp = {};

    for (var key in worksheet) {
        var cell = worksheet[key],
            match = /([A-Z]+)(\d+)/.exec(key);

        if (!match) {
            continue;
        }

        var row = match[2]; // 1234
        var col = match[1]; // ABCD

        if (col >= startColumn) {
            if (row == headRow) {
                tmp[key] = cell.v;
                headers.push(tmp);
            }
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
var getHeadersValue = function(_headers, headRow, _key) {
    var res = null;
    var key = _key + '' + headRow;

    _headers.forEach(function(item) {
        if (item.hasOwnProperty(key)) {
            res = item[key];
        }
    })

    return res;
}

var createItem = function(headers) {
    var res = {};
    // headers.forEach(function(item, index) {
    //     log(item + ' - ' + index);
    //     obj[item] = null;
    // });

    for (var obj in headers) {
        if (headers.hasOwnProperty(obj)) {
            for (var prop in headers[obj]) {
                if (headers[obj].hasOwnProperty(prop)) {
                    // log(prop + ':' + headers[obj][prop]);
                    res[headers[obj][prop]] = null;
                }
            }
        }
    }

    return res;
};

/**
 * excel filename or workbook to json
 * @param fileName
 * @param headRow
 * @param valueRow
 * @returns {Array} json
 */
var toJson = function(fileName, headRow, valueRow, startColumn) {
    var workbook;

    if (typeof fileName === 'string') {
        workbook = XLSX.readFile(fileName);
    } else {
        workbook = fileName;
    }

    var worksheet = workbook.Sheets[workbook.SheetNames[0]],
        json = [],
        namemap = {},
        curRow = 0,
        value,
        row,
        col,
        cell,
        match,
        index,
        headers = getHeaders(worksheet, headRow, startColumn);

    for (var key in worksheet) {
        if (worksheet.hasOwnProperty(key)) {
            cell = worksheet[key];
            match = /([A-Z]+)(\d+)/.exec(key);

            if (!match) {
                continue;
            }

            row = match[2]; // 1234
            col = match[1]; // ABCD

            value = cell.v;

            if (col >= startColumn) {
                if (row < valueRow) {
                    //continue;
                } else {
                    // check if we have a new line
                    if (row > curRow) {
                        curRow = row;
                        namemap = createItem(headers);
                    }

                    namemap[getHeadersValue(headers, headRow, col)] = value;

                    json[row - headRow - 1] = JSON.parse(JSON.stringify(namemap));
                }
            }
        }
    }
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
        var workbook = XLSX.read(bString, {type: "binary", sheetStubs: true, cellHTML: true});
        file.contents = new Buffer(JSON.stringify(toJson(workbook, options.headRow || 1, options.valueRowStart || 2, options.startColumn || 'A')));

        log("Convert file: " + file.path);

        this.push(file);
        cb();
    });
};
