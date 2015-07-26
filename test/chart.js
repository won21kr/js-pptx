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
        var slide2 = pptx.addSlide("slideLayout1");
        slide1.addChart(function (err, chart) {
          console.log("DONE ADDING CHART1");
          fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
            if (err) throw err;
            console.log("open " + OUTFILE)
            done();
          });
        });
      });
    });
  });
});
