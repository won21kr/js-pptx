"use strict";

var assert = require('assert');
var fs = require("fs");
var PPTX = require('../lib/pptx');
var Query = require('query');


var INFILE = './test/files/parts3.pptx';
var OUTFILE = './test/files/parts3-a.pptx';

describe('PPTX', function () {

  it('can read, modify, write and read', function (done) {
    fs.readFile(INFILE, function (err, data) {
      if (err) throw err;
      var pptx = new PPTX.Presentation();
      pptx.load(data, function (err) {

        var slide1 = pptx.getSlide('slide1');
        var shapes = slide1.getShapes();

        shapes[3].shapeProperties().x(PPTX.emu.inch(1));
        shapes[3].shapeProperties().y(PPTX.emu.inch(1));

        shapes[3].shapeProperties().cx(PPTX.emu.inch(2));
        shapes[3].shapeProperties().cy(PPTX.emu.inch(0.75));
        shapes[3].shapeProperties().prstGeom('trapezoid');
        shapes[3].text("Now it's a trapezoid")

        fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
          if (err) throw err;

          fs.readFile(OUTFILE, function (err, data) {
            var check = new PPTX.Presentation();
            check.load(data, function (err) {

              var props = check.getSlide('slide1').getShapes()[3].shapeProperties().toJSON();
              assert.deepEqual(props, { x: '914400',
                y: '914400',
                cx: '1828800',
                cy: '685800',
                prstGeom: 'trapezoid' });
            });
            done();
          });
        });
      });
    });
  });
});
