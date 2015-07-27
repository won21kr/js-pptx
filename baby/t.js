"use strict";


var JSZip = require('jszip');
var fs = require("fs");

var TESTFILE = '/tmp/chart.pptx';
var REFERENCE = './lab/chart/chart.pptx';
var OUTFILE = '/tmp/chart2.pptx';

function $(xml) {
  return cheerio.load(xml, {xmlMode: true})
}


var testZip = new JSZip(fs.readFileSync(TESTFILE));
var refZip = new JSZip(fs.readFileSync(REFERENCE));
var outZip = new JSZip();

//console.log(Object.keys(refZip.files));
//process.exit()

// copy all the files from testzip to outzip
// copy selected files from refzip to outzuip
// save outzip

Object.keys(testZip.files).forEach(function (key) {
  outZip.file(key, testZip.file(key).asArrayBuffer());
});

var keys = [
//  '[Content_Types].xml',
//  '_rels/.rels',
//  'ppt/slides/_rels/slide2.xml.rels',
//  'ppt/slides/_rels/slide1.xml.rels',
//  'ppt/_rels/presentation.xml.rels',
//  'ppt/presentation.xml',
//  'ppt/slides/slide2.xml',
//  'ppt/slides/slide1.xml',
//  'ppt/slideLayouts/_rels/slideLayout8.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout9.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout5.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout6.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout4.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout2.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout1.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout11.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout10.xml.rels',
//  'ppt/slideMasters/_rels/slideMaster1.xml.rels',
//  'ppt/slideLayouts/_rels/slideLayout3.xml.rels',
//  'ppt/slideLayouts/slideLayout10.xml',
//  'ppt/slideLayouts/slideLayout9.xml',
//  'ppt/slideLayouts/slideLayout2.xml',
//  'ppt/slideLayouts/slideLayout1.xml',
//  'ppt/slideMasters/slideMaster1.xml',
//  'ppt/slideLayouts/_rels/slideLayout7.xml.rels',
//  'ppt/slideLayouts/slideLayout3.xml',
//  'ppt/slideLayouts/slideLayout4.xml',
//  'ppt/slideLayouts/slideLayout5.xml',
//  'ppt/slideLayouts/slideLayout6.xml',
//  'ppt/slideLayouts/slideLayout7.xml',
//  'ppt/slideLayouts/slideLayout8.xml',
//  'ppt/slideLayouts/slideLayout11.xml',
//  'ppt/embeddings/Microsoft_Excel_Sheet2.xlsx',
//  'ppt/charts/_rels/chart2.xml.rels',
//  'ppt/theme/theme1.xml',
//  'ppt/charts/chart1.xml',
//  'ppt/embeddings/Microsoft_Excel_Sheet1.xlsx',
//  'ppt/charts/chart2.xml',
//  'docProps/thumbnail.jpeg',
//  'ppt/charts/_rels/chart1.xml.rels',
//  'ppt/viewProps.xml',
//  'ppt/tableStyles.xml',
//  'ppt/presProps.xml',
//  'docProps/app.xml',
//  'docProps/core.xml',
//  'ppt/printerSettings/printerSettings1.bin'
]

keys.forEach(function (key) {
  outZip.file(key, refZip.file(key).asArrayBuffer());
});

var buffer = outZip.generate({type: "nodebuffer"});

fs.writeFile(OUTFILE, buffer, function (err) {
  if (err) throw err;
  console.log("open "+OUTFILE)
});
