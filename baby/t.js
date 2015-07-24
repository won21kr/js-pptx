"use strict";


var JSZip = require('jszip');
var fs = require("fs");

var TESTFILE = './test/files/parts3-a.pptx';
var REFERENCE = './lab/parts3-b/parts3-b.pptx';
var OUTFILE = '/tmp/parts-3c.pptx';

function $(xml) {
  return cheerio.load(xml, {xmlMode: true})
}


var testZip = new JSZip(fs.readFileSync(TESTFILE));
var refZip = new JSZip(fs.readFileSync(REFERENCE));
var outZip = new JSZip();


// copy all the files from testzip to outzip
// copy selected files from refzip to outzuip
// save outzip

Object.keys(testZip.files).forEach(function (key) {
  outZip.file(key, testZip.file(key).asText());
});

var keys = [
//  "[Content_Types].xml",
//  "ppt/_rels/presentation.xml.rels",
//  "ppt/presentation.xml",
//  "ppt/slideLayouts/slideLayout11.xml",
//  "ppt/slideLayouts/slideLayout10.xml",
//  "ppt/slideLayouts/slideLayout9.xml",
//  "ppt/slideLayouts/slideLayout2.xml",
//  "ppt/slideLayouts/slideLayout1.xml",
//  "ppt/slideMasters/slideMaster1.xml",
//  "ppt/slideLayouts/slideLayout3.xml",
//  "ppt/slideLayouts/slideLayout4.xml",
//  "ppt/slideLayouts/slideLayout5.xml",
//  "ppt/slideLayouts/slideLayout6.xml",
//  "ppt/slideLayouts/slideLayout7.xml",
//  "ppt/slideLayouts/slideLayout8.xml",
//  "docProps/thumbnail.jpeg",
//  "ppt/viewProps.xml",
//  "docProps/core.xml",
//  "ppt/slides/slide3.xml"
]

keys.forEach(function (key) {
  outZip.file(key, refZip.file(key).asText());
});

var buffer = outZip.generate({type: "nodebuffer"});

fs.writeFile(OUTFILE, buffer, function (err) {
  if (err) throw err;
});
