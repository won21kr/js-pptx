"use strict";

var assert = require('assert');
var fs = require("fs");
var JSZip = require('jszip')

var INFILE = './test/files/minimal.pptx';
var STORE = './test/files/minimal.json';
var OUTFILE = './test/files/minimal-copy.pptx';

var zip1 = new JSZip(fs.readFileSync(INFILE));
var copy = {};

Object.keys(zip1.files).forEach(function (key) {
  copy[key] = zip1.file(key).asText();
});

fs.writeFileSync(STORE, JSON.stringify(copy, null,4))

var json = fs.readFileSync(STORE, 'utf8');
var obj = JSON.parse(json);

var zip2 = new JSZip();
for (var key in obj) {
  zip2.file(key, obj[key]);
}

var buffer = zip2.generate({type:"nodebuffer", compression: 'DEFLATE'});

fs.writeFile(OUTFILE, buffer, function(err) {
  if (err) throw err;
});

var zip3 = new JSZip();
zip3.file('json', JSON.stringify(copy, null,4))
var buffer3 = zip3.generate({type:"nodebuffer", compression: 'DEFLATE'});

fs.writeFile('./test/files/minimal.json.jar', buffer3, function(err) {
  if (err) throw err;
});