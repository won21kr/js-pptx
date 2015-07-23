"use strict";

// open an existing doc, unzip it, zip it, and save it as new name.  See if it opens.


var JSZip = require('jszip');
var fs = require("fs");

var INFILE = './lab/out/out.pptx';
var OUTFILE = './lab/baby.pptx';

var STYLEFILE = './lab/ex1/ex1.pptx';

// read a zip file
fs.readFile(INFILE, function(err, data) {
  if (err) throw err;
  var zip = new JSZip(data);


  fs.readFile(STYLEFILE, function(err, data2) {
    var zip_theme = new JSZip(data2)
    var theme_xml = zip_theme.file('ppt/theme/theme1.xml').asText();
    zip.file('ppt/theme/theme1.xml', theme_xml);

    var buffer = zip.generate({type:"nodebuffer"});

    fs.writeFile(OUTFILE, buffer, function(err) {
      if (err) throw err;
    });
  });

});