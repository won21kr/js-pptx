"use strict";

var fs = require("fs");
var PPTX = require('../lib/pptx');
var Query = require('query');

Array.prototype.get = function (filter) {
  return Query.query(this, filter);
}

//console.log([{a: 1}, {a:2}, {a:3}].query({a: { $gt: 1}}))


var INFILE = './lab/parts3/parts3.pptx';
var OUTFILE = './lab/parts2/parts2.pptx';


fs.readFile(INFILE, function (err, data) {
  if (err) throw err;
  var pptx = new PPTX.Presentation();
  pptx.load(data, function (err) {


    var slide1 = pptx.getSlide('slide1');
    var shapes = slide1.getShapes();
    console.log(shapes[3].text())

    console.log(JSON.stringify(shapes[3].shapeProperties().toJSON(),null,4))

    shapes[3].shapeProperties().x(PPTX.emu.inch(1))
    shapes[3].shapeProperties().y(PPTX.emu.inch(1))

    shapes[3].shapeProperties().cx(PPTX.emu.inch(2))
    shapes[3].shapeProperties().cy(PPTX.emu.inch(0.75))

    shapes[3].shapeProperties().prstGeom('trapezoid');

    console.log(JSON.stringify(shapes[3].shapeProperties().toJSON(),null,4))

    //fs.writeFile('./lab/parts2/parts2.json', JSON.stringify(pptx, null,4), 'utf8');
    fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
      if (err) throw err;
    });
  });
});

