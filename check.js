#! /usr/bin/env node

var inquirer = require("inquirer");
var XLSX     = require('xlsx');
var url      = require('url');
var fs       = require('fs');
var request  = require('request');
var colors   = require('colors');

process.stdin.setMaxListeners(0);

console.log( 'XLS Redirect Checker'.underline.rainbow );

var questions = [
    { type: "input", name: "sourcePath", message: "Full path to the sheet?" },
    { type: "input", name: "orginalStartCell", message: "What is the first cell containing the original urls?", default: "A2" },
    { type: "input", name: "redirectStartCell", message: "What is the first cell containing the redirect urls?", default: "B2" },
    { type: "confirm", name: "basicAuth", message: "Do these links require basic auth?" },
    { type: "input", name: "basicAuthUser", message: "Basic auth user:", when: function( response ) { return response.basicAuth; } },
    { type: "input", name: "basicAuthPwd", message: "Basic auth password:", when: function( response ) { return response.basicAuth; } }
];

inquirer.prompt( questions, function( answers ) {
    var workbook     = XLSX.readFile( answers.sourcePath );
    var sheetName    = workbook.SheetNames[ 0 ]; // Always use the first sheet
    var sheet        = workbook.Sheets[ sheetName ];
    var originalUrls  = { column: answers.orginalStartCell.substr( 0, 1 ), row: parseInt( answers.orginalStartCell.substr( 1, 1 ) ) };
    var redirectUrls = { column: answers.redirectStartCell.substr( 0, 1 ), row: parseInt( answers.redirectStartCell.substr( 1, 1 ) ) };
    var outputString = "";
    var passed = 0;
    var failed = 0;
    var total  = 0;
    var start = new Date().getTime();

    // Build any necessary request headers
    var headers = {};
    if ( answers.basicAuth ) {
        headers["Authorization"] =  "Basic " + new Buffer( answers.basicAuthUser + ":" + answers.basicAuthPwd ).toString( "base64" );
    }

    // Get the first original and redirect url
    var currentOriginal = sheet[ originalUrls.column + originalUrls.row ];
    var currentRedirect = sheet[ redirectUrls.column + redirectUrls.row ];

    check();

    // Loops over the sheet until there are no urls left and create an output string
    function check() {

        if ( currentOriginal && currentRedirect.v ) {
            var r = request.get( {
                    url: currentOriginal.v,
                    //maxRedirects: 1,
                    headers: headers
                } , function (err, res, body) {
                    //var removeSlash = currentRedirect.v.replace(/\/$/, "");
                    if ( res && decodeURIComponent( res.request.uri.href.replace(/\/$/, "") ) == currentRedirect.v.replace(/\/$/, "") ) {
                        passed++;
                        console.log( originalUrls.column + originalUrls.row + ' passed' );
                    } else {
                        failed++;

                        outputString += originalUrls.column + originalUrls.row + "\n";
                        outputString += "Visited: " + currentOriginal.v + "\n";
                        outputString += "Expected: " + currentRedirect.v + "\n";
                        if ( res ) {
                            outputString += "Got: " + res.request.uri.href + "\n\n";
                        } else if ( err ) {
                            outputString += "Got: there was an error" + err.code + "\n\n";
                        } else {
                            outputString += "Got: Something broke \n\n";
                        }

                        console.log( originalUrls.column + originalUrls.row + ' failed' );
                    }

                    total++;

                    console.log( colors.green( "Passed: " + passed ) + colors.red( " Failed: " + failed ) + colors.cyan( " Total: " + total ) );

                    // Get the next urls
                    originalUrls.row++;
                    redirectUrls.row++;

                    currentOriginal = sheet[ originalUrls.column + originalUrls.row ];
                    currentRedirect = sheet[ redirectUrls.column + redirectUrls.row ];

                    check();

            } );
        } else {

            var end = new Date().getTime();
            var time = ( end - start ) / 1000;

            console.log( colors.cyan( "\n\nFinished in: " + time + "s" ) );
            console.log( colors.green( "Passed: " + passed ) );
            console.log( colors.red( "Failed: " + failed ) );
            console.log( "Total: " + total );

            fs.writeFile( "failed.txt", outputString, function( err ) {
                if (err) throw err;
                console.log( "Created failed file" );
            } );
        }

    }

});
