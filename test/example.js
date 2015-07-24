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

        var slide3 = pptx.addSlide("slideLayout1");
        var slide4 = pptx.addSlide("slideLayout2");

        var slide1 = pptx.getSlide('slide1');
        var shapes = slide1.getShapes()
        shapes[3]
            .text("Now it's a trapezoid")
            .shapeProperties()
            .x(PPTX.emu.inch(1))
            .y(PPTX.emu.inch(1))
            .cx(PPTX.emu.inch(2))
            .cy(PPTX.emu.inch(0.75))
            .prstGeom('trapezoid');


        var triangle = slide1.addShape()
            .text("Triangle")
            .shapeProperties()
            .x(PPTX.emu.inch(2))
            .y(PPTX.emu.inch(2))
            .cx(PPTX.emu.inch(2))
            .cy(PPTX.emu.inch(2))
            .prstGeom('triangle');

        for (var i= 0; i<20; i++) {
          slide1.addShape()
              .text(""+i)
              .shapeProperties()
              .x(PPTX.emu.inch((Math.random()*10)))
              .y(PPTX.emu.inch((Math.random()*6)))
              .cx(PPTX.emu.inch(1))
              .cy(PPTX.emu.inch(1))
              .prstGeom('ellipse');
        }

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
