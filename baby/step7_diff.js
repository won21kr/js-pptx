"use strict";


var JSZip = require('jszip');
var fs = require("fs");
var cheerio = require('cheerio');
var xmldoc = require('xmldoc');
var pd = require('pretty-data').pd;
var pretty = function(xml) { return pd.xml(xml); }

var FILE1 = './lab/parts3/parts3.pptx';
var FILE2 = './lab/parts2/parts2.pptx';


var zip1 = new JSZip(fs.readFileSync(FILE1));
var zip2 = new JSZip(fs.readFileSync(FILE2));

Object.keys(zip2.files).forEach(function(key) {

  var str1 = zip1.file(key).asText().replace(/\n|\s/ig, '');
  var str2 = zip2.file(key).asText().replace(/\n|\s/ig, '');


  if (str1 != str2) {
    console.log(key, str1.length, str2.length, str1.length== str2.length)
}

})
