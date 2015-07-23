"use strict";

var fs = require("fs");
var PPTX = require('../lib/pptx');

var INFILE = './lab/parts3/parts3.pptx';
var OUTFILE = './lab/parts2/parts2.pptx';


fs.readFile(INFILE, function(err, data) {
  if (err) throw err;
  var pptx = new PPTX();

  pptx.load(data, function (err) {

    fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
      if (err) throw err;
    });
  });


});




