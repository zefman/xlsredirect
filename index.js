#! /usr/bin/env node

var inquirer = require("inquirer");
var XLSX     = require('xlsx');
var url      = require('url');
var fs       = require('fs');
var colors   = require('colors');

console.log( 'XLS Redirect Maker\n\n'.cyan.bold.underline );

var questions = [
    { type: "input", name: "sourcePath", message: "Full path to the spreadsheet file?" },
    { type: "input", name: "orginalStartCell", message: "What is the first cell containing the original urls?", default: "A2" },
    { type: "input", name: "redirectStartCell", message: "What is the first cell containing the redirect urls?", default: "B2" },
    { type: "input", name: "outputPath", message: "Where do you want the output file?", default: "output.txt" },
    { type: "input", name: "rewriteFlags", message: "Specify reqrite flags:", default: "[R=301,L,QSD]" }
];

inquirer.prompt( questions, function( answers ) {
    var workbook     = XLSX.readFile( answers.sourcePath );
    var sheetName    = workbook.SheetNames[ 0 ]; // Always use the first sheet
    var sheet        = workbook.Sheets[ sheetName ];
    var originalUrls  = { column: answers.orginalStartCell.substr( 0, 1 ), row: parseInt( answers.orginalStartCell.substr( 1, 1 ) ) };
    var redirectUrls = { column: answers.redirectStartCell.substr( 0, 1 ), row: parseInt( answers.redirectStartCell.substr( 1, 1 ) ) };
    var outputString = "";

    // Get the first original and redirect url
    var currentOriginal = sheet[ originalUrls.column + originalUrls.row ];
    var currentRedirect = sheet[ redirectUrls.column + redirectUrls.row ];

    // Loops over the sheet until there are no urls left and create an output string
    while ( currentOriginal && currentRedirect.v ) {
        var parsedOriginal = url.parse( currentOriginal.v );
        var parsedRedirect = url.parse( currentRedirect.v );

        if ( parsedOriginal.pathname.replace( /^\//, "" ) != "" ) {
            // If there is a query in the original url we need to add a RewriteCond
            if ( parsedOriginal.query ) {
                outputString += "RewriteCond %{QUERY_STRING} ^" + parsedOriginal.query + "$\n";
            }

            // Create the rewrite rule
            outputString += "RewriteRule ^" + parsedOriginal.pathname.replace( /^\//, "" ) + "$ " + decodeURIComponent( currentRedirect.v );

            // If the original url has a query string and the redirect doesn't we need to add a
            // ? to the redirect url to prevent the query being passed forward
            if ( parsedOriginal.query && !parsedRedirect.query ) {
                // For some reason, to make sure the query doesn't pass on parameters, the redirect
                // url must end in a slash followed by the ?
                if ( outputString.substr(-1) != '/' ) {
                    outputString += '/';
                }

                outputString += "? " + answers.rewriteFlags + "\n";
            } else {
                outputString += " " + answers.rewriteFlags + "\n";
            }

        }

        // Get the next urls
        originalUrls.row++;
        redirectUrls.row++;

        currentOriginal = sheet[ originalUrls.column + originalUrls.row ];
        currentRedirect = sheet[ redirectUrls.column + redirectUrls.row ];
    }

    fs.writeFile( answers.outputPath, outputString, function( err ) {
        if (err) throw err;
        console.log( colors.green( "Created output redirect file" ) );
    } );
});
