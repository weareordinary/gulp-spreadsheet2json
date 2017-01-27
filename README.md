# gulp-spreadsheet2json
> Export Excel (XLSX/XLS/ODS) with headers to json (only first spreadsheet in file).


[![NPM](https://nodei.co/npm/gulp-spreadsheet2json.png?downloads=true&downloadRank=true&stars=true)](https://nodei.co/npm/gulp-spreadsheet2json/)

# ATTENTION!

> I found a small bug with OpenOffice `ODT` files. If table has many columns, `XLSX.js` doesn't read all columns, but same table in `XLS` format work stabil!

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
        xls2json = require('gulp-spreadsheet2json');

    var spreadsheets = [
            'config/**.xlsx',
            'config/**.xls',
            'config/**.ods'
        ];

    gulp.task('parse:spreadsheet', function() {
        gulp.src(spreadsheets)
            .pipe(xls2json({
                headRow: 1,
                valueRowStart: 2,
                trace: false
            }))
            .pipe(rename(function(path) {
                path.extname = ".json";
            }))
            .pipe(gulp.dest('build'));
    });

    gulp.task('default', ['parse:spreadsheet']);

}());
```

> Input

color_name	| R	| G	| B
--- | :---: | :---: | :---:
illuminant	| 255	| 255	| 255
dark skin	| 107	| 83	| 70
light skin	| 182	| 147	| 128

> Output

```
[
    {
        "color_name": "illuminant",
        "R": 255.00000000002,
        "G": 254.999999999984,
        "B": 254.999999999997
    },
    {
        "color_name": "dark skin",
        "R": 106.732127788008,
        "G": 82.6909148074604,
        "B": 70.141586954399
    },
    {
        "color_name": "light skin",
        "R": 182.148358411673,
        "G": 147.240481111174,
        "B": 128.262943740921
    }
]
```


## API

### spreadsheet2json([options])

Name | Type | Default | Description
--- | :---: | :---: | ---
headRow | `number` | `1` | The row number of head. (Start from 1)
valueRowStart | `number` | `2` | The start row number of values. (Start from 1)
startColumn | `string` | `A` | The start column Char of values. (Start from A)
trace | `Boolean` | `false` | Whether to log each file path while convert success.


## TODO

* add option for workbook.Sheets index
* add limit for rows

## License
MIT
