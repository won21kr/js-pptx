"use strict";

var assert = require('assert');
var fs = require("fs");
var PPTX = require('../lib/pptx');
var xml2js = require('xml2js');
var xmlbuilder = require('xmlbuilder')


var INFILE = './lab/chart-null/chart-null.pptx';
var OUTFILE = './lab/chart-one/chart-one.pptx';

describe('PPTX', function () {

  it('can read, modify, write and read', function (done) {
    fs.readFile(INFILE, function (err, data) {
      if (err) throw err;
      var pptx = new PPTX.Presentation();


      pptx.load(data, function (err) {

        var slide1 = pptx.getSlide('slide1');
//        var shapes = slide1.getShapes()
//        shapes[3]
//            .text("Now it's a trapezoid")
//            .shapeProperties()
//            .x(PPTX.emu.inch(1))
//            .y(PPTX.emu.inch(1))
//            .cx(PPTX.emu.inch(2))
//            .cy(PPTX.emu.inch(0.75))
//            .prstGeom('trapezoid');
//
//
//        var triangle = slide1.addShape()
//            .text("Triangle")
//            .shapeProperties()
//            .x(PPTX.emu.inch(2))
//            .y(PPTX.emu.inch(2))
//            .cx(PPTX.emu.inch(2))
//            .cy(PPTX.emu.inch(2))
//            .prstGeom('triangle');

//        for (var i= 0; i<20; i++) {
//          slide1.addShape()
//              .text(""+i)
//              .shapeProperties()
//              .x(PPTX.emu.inch((Math.random()*10)))
//              .y(PPTX.emu.inch((Math.random()*6)))
//              .cx(PPTX.emu.inch(1))
//              .cy(PPTX.emu.inch(1))
//              .prstGeom('ellipse');
//        }
//
//        fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
//            if (err) throw err;
//          console.log(OUTFILE);
//          done()  ;
//        });

        var chart = slide1.addChart(function (err, chart) {
          console.log("DONE ADDING CHART");

//          console.log("############ " + pptx.content["docProps/thumbnail.jpeg"].length)
          fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
            if (err) throw err;
            console.log("open "+OUTFILE)
            done();
          });
        });
      });
    });
  });
});
