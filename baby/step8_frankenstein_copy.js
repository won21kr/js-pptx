"use strict";


var JSZip = require('jszip');
var fs = require("fs");
var cheerio = require('cheerio');
var xmldoc = require('xmldoc');
var pd = require('pretty-data').pd;
var pretty = function(xml) { return pd.xml(xml); }

var INFILE = './lab/parts3/parts3.pptx';
var OUTFILE = './lab/parts2/parts2.pptx';

function $(xml) { return cheerio.load(xml, {xmlMode: true})}

fs.readFile(INFILE, function(err, data) {
  if (err) throw err;
  var zip1 = new JSZip(data);

  var zip2 = new JSZip();


  Object.keys(zip1.files).forEach(function(key) {
    console.log(key)
    zip2.file(key, zip1.file(key).asText() )  ;
  });


  var buffer = zip2.generate({type:"nodebuffer"});

  fs.writeFile(OUTFILE, buffer, function(err) {
    if (err) throw err;
  });
//

});
