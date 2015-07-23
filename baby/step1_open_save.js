"use strict";

// open an existing doc, unzip it, zip it, and save it as new name.  See if it opens.


var JSZip = require('jszip');
var fs = require("fs");

var INFILE = './lab/out/out.pptx';
var OUTFILE = './lab/baby.pptx';


// read a zip file
fs.readFile(INFILE, function(err, data) {
  if (err) throw err;
  var zip = new JSZip(data);
  var buffer = zip.generate({type:"nodebuffer"});

  fs.writeFile(OUTFILE, buffer, function(err) {
    if (err) throw err;
  });
});