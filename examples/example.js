"use strict";

var fs = require("fs");
var PPTX = require('..');


var INFILE = './test/files/minimal.pptx'; // a blank PPTX file with my layouts, themes, masters.
var OUTFILE = '/tmp/example.pptx';

fs.readFile(INFILE, function (err, data) {
  if (err) throw err;
  var pptx = new PPTX.Presentation();
  pptx.load(data, function (err) {
    var slide1 = pptx.getSlide('slide1');

    var slide2 = pptx.addSlide("slideLayout3"); // section divider
    var slide3 = pptx.addSlide("slideLayout6"); // title only


    var triangle = slide1.addShape()
        .text("Triangle")
        .shapeProperties()
        .x(PPTX.emu.inch(2))
        .y(PPTX.emu.inch(2))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(2))
        .prstGeom('triangle');

    var triangle = slide1.addShape()
        .text("Ellipse")
        .shapeProperties()
        .x(PPTX.emu.inch(4))
        .y(PPTX.emu.inch(4))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(1))
        .prstGeom('ellipse');

    for (var i = 0; i < 20; i++) {
      slide2.addShape()
          .text("" + i)
          .shapeProperties()
          .x(PPTX.emu.inch((Math.random() * 10)))
          .y(PPTX.emu.inch((Math.random() * 6)))
          .cx(PPTX.emu.inch(1))
          .cy(PPTX.emu.inch(1))
          .prstGeom('ellipse');
    }

    slide1.getShapes()[3]
        .text("Now it's a trapezoid")
        .shapeProperties()
        .x(PPTX.emu.inch(1))
        .y(PPTX.emu.inch(1))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(0.75))
        .prstGeom('trapezoid');

    var chart = slide3.addChart(barChart, function (err, chart) {

      fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
        if (err) throw err;
        console.log("open " + OUTFILE)
      });
    });
  });
})
;

var barChart = {
  title: 'Sample bar chart',
  renderType: 'bar',
  data: [
    {
      name: 'Series 1',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [4.3, 2.5, 3.5, 4.5]
    },
    {
      name: 'Series 2',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [2.4, 4.4, 1.8, 2.8]
    },
    {
      name: 'Series 3',
      labels: ['Category 1', 'Category 2', 'Category 3', 'Category 4'],
      values: [2.0, 2.0, 3.0, 5.0]
    }
  ]
}