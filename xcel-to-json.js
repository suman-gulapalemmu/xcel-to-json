'use strict';

/**
 * Module dependencies.
 */
var fs = require('fs');
var xlsx = require('xlsx');
var cvcsv = require('csv');

exports = module.exports = XLS_TO_JSON;

function XLS_TO_JSON(config, callback) {
    if (!config.input) {
        console.error('You miss a input file');
        process.exit(1);
    }

    var cv = new CV(config, callback);
}

function CV(config, callback) {
    var wb = this.load_wb(config.input);
    var ws = this.load_ws(wb, config.sheet);
    var csv = this.get_csv(ws);
    this.convertToJSON(csv, config.output, callback);
}

CV.prototype.load_wb = function (input) {
    return xlsx.readFile(input);
};

CV.prototype.load_ws = function (wb, target_sheet) {
    return wb.Sheets[target_sheet ? target_sheet : wb.SheetNames[0]];
};

CV.prototype.get_csv = function (ws) {
    return xlsx.utils.make_csv(ws);
};

CV.prototype.convertToJSON = function (csv, output, callback) {
    var headerArray = [];
    var outputObj = [];

    cvcsv()
        .from.string(csv)
        .transform(function (row) {
            row.unshift(row.pop());
            return row;
        })
        .on('record', function (row, index) {
            if (index === 0) {
                headerArray = row;
            } else {
                var rowObj = {};

                headerArray.forEach(function (column, index) {
                    if (column.trim().indexOf('.') !== -1) {
                        // Create Deep Link Object
                        createDeepLinkObject(column, row[index].trim(), rowObj);
                    } else {
                        // No nested objects
                        rowObj[column.trim()] = row[index].trim();
                    }
                });

                outputObj.push(rowObj);
            }
        })
        .on('end', function (count) {
            // when writing to a file, use the 'close' event
            // the 'end' event may fire before the file has been written

            if (output && output !== null) {
                var stream = fs.createWriteStream(output, {flags: 'w'});
                stream.write(JSON.stringify(outputObj));
            }
            callback(null, outputObj);
        })
        .on('error', function (error) {
            callback(error, null);
        });

    function createDeepLinkObject(headerColumnElement, rowValue, object) {
        var headerParts = headerColumnElement.trim().split('.');

        if (headerParts.length === 1) {
            object[headerParts[0]] = rowValue;
        }
        else {
            if (!object[headerParts[0]]) object[headerParts[0]] = {};
            object = object[headerParts.shift()];

            createDeepLinkObject(headerParts.join('.'), rowValue, object);
        }
    }
};
