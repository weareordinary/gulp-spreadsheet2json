# gulp-spreadsheet2json
> Export Excel (XLSX/XLS/ODS) with headers to json (only first spreadsheet in file).


[![NPM](https://nodei.co/npm/gulp-spreadsheet2json.png?downloads=true&downloadRank=true&stars=true)](https://nodei.co/npm/gulp-spreadsheet2json/)

## Usage
First, install `gulp-spreadsheet2json` as a development dependency:

```shell
npm install --save-dev gulp-spreadsheet2json
```

Then, add it to your `gulpfile.js`:

```javascript
(function() {
    'use strict';

    var gulp = require('gulp'),
        rename = require("gulp-rename"),
        del = require("del"),
        xls2json = require('gulp-spreadsheet2json'),
        spreadsheets = [
            'config/**.xlsx',
            'config/**.xls',
            'config/**.ods'
        ];

    gulp.task('clean', function() {
        del('build');
    });
    
    gulp.task('copy', ['clean'], function() {
        gulp.src(spreadsheets)
            .pipe(xls2json({
                headRow: 1,
                valueRowStart: 2,
                startColumn: 'C',
                trace: false
            }))
            .pipe(rename(function(path) {
                path.extname = ".json";
            }))
            .pipe(gulp.dest('build'));
    });

}());
```

## API

### spreadsheet2json([options])

Name | Type | Default | Description
--- | :---: | :---: | ---
headRow | `number` | `1` | The row number of head. (Start from 1)
valueRowStart | `number` | `3` | The start row number of values. (Start from 1)
startColumn | `number` | `A` | The start column Char of values. (Start from A)
trace | `Boolean` | `false` | Whether to log each file path while convert success.


## TODO

* add options for workbook.Sheets index
* add limit for rows

## License
MIT
