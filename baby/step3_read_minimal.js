"use strict";


var JSZip = require('jszip');
var fs = require("fs");

var INFILE = './lab/minimal/minimal.pptx';

// read a zip file
fs.readFile(INFILE, function(err, data) {
  if (err) throw err;
  var zip = new JSZip(data);

  // ok, yeah, no wow, we've got a pptx file open.  Can I modify it somehow

  console.log(Object.keys(zip.files))

  console.log(zip.file('ppt/slides/slide1.xml').asText())

});



var x = ["[Content_Types].xml",
  "_rels/.rels",
  "ppt/slides/_rels/slide1.xml.rels",
  "ppt/_rels/presentation.xml.rels",
  "ppt/presentation.xml",
  "ppt/slides/slide1.xml",
  "ppt/slideLayouts/_rels/slideLayout6.xml.rels",
  "ppt/slideMasters/_rels/slideMaster1.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout8.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout9.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout11.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout10.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout7.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout2.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout3.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout4.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout5.xml.rels",
  "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
  "ppt/slideLayouts/slideLayout11.xml",
  "ppt/slideMasters/slideMaster1.xml",
  "ppt/slideLayouts/slideLayout1.xml",
  "ppt/slideLayouts/slideLayout2.xml",
  "ppt/slideLayouts/slideLayout3.xml",
  "ppt/slideLayouts/slideLayout10.xml",
  "ppt/slideLayouts/slideLayout4.xml",
  "ppt/slideLayouts/slideLayout6.xml",
  "ppt/slideLayouts/slideLayout5.xml",
  "ppt/slideLayouts/slideLayout9.xml",
  "ppt/slideLayouts/slideLayout7.xml",
  "ppt/slideLayouts/slideLayout8.xml",
  " docProps/thumbnail.jpeg",
  "ppt/theme/theme1.xml",
  "ppt/viewProps.xml",
  "ppt/tableStyles.xml",
  "ppt/presProps.xml",
  "docProps/app.xml",
  "docProps/core.xml",
  "ppt/printerSettings/printerSettings1.bin"
];
