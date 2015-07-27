"use strict";

var assert = require('assert');
var fs = require("fs");
var PPTX = require('../lib/pptx');
var xml2js = require('xml2js');
var xmlbuilder = require('xmlbuilder')


var INFILE = './lab/chart-null/chart-null.pptx';
var OUTFILE = '/tmp/chart.pptx';

describe('PPTX', function () {

  it('can read, modify, write and read', function (done) {
    fs.readFile(INFILE, function (err, data) {
      if (err) throw err;
      var pptx = new PPTX.Presentation();


      pptx.load(data, function (err) {

        var slide1 = pptx.getSlide('slide1');
        var slide2 = pptx.addSlide("slideLayout1");
        slide1.addChart(barChart, function (err, chart) {
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


var barChart = {
  title: 'Sample bar chart',
  renderType: 'bar',
  xmlOptions: {
    "c:title": {
      "c:tx": {
        "c:rich": {
          "a:p": {
            "a:r": {
              "a:t": "Override title via XML"
            }
          }
        }
      }
    }
  },
  data: [
    {
      name: 'europe',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [2.5, 2.6, 2.8],
      color: 'ff0000'
    },
    {
      name: 'namerica',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [2.5, 2.7, 2.9],
      color: '00ff00'
    },
    {
      name: 'asia',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [2.1, 2.2, 2.4],
      color: '0000ff'
    },
    {
      name: 'lamerica',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [0.3, 0.3, 0.3],
      color: 'ffff00'
    },
    {
      name: 'meast',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [0.2, 0.3, 0.3],
      color: 'ff00ff'
    },
    {
      name: 'africa',
      labels: ['Y2003', 'Y2004', 'Y2005'],
      values: [0.1, 0.1, 0.1],
      color: '00ffff'
    }

  ]
};