"use strict";


var JSZip = require('jszip');
var fs = require("fs");
var cheerio = require('cheerio');
var pd = require('pretty-data').pd;

var INFILE = './lab/parts/parts.pptx';
var OUTFILE = './lab/parts/parts2.pptx';


var $ = cheerio.load

// read a zip file
fs.readFile(INFILE, function(err, data) {
  if (err) throw err;
  var zip = new JSZip(data);
  var str0 = zip.file('ppt/slides/slide1.xml').asText();
  var $slide1 = cheerio.load(zip.file('ppt/slides/slide1.xml').asText(), {xmlMode: true});

  var str1 = $slide1.xml();
  console.log([str0.length, str1.length, str0==str1])
  var xml = $slide1.xml();

//  console.log(pretty);
  $slide1('p\\:sp').each(function(i, elem) {

    console.log('---------------------------------------')
    $slide1('a\\:t', elem).text("This is different")
    var a = $slide1('a\\:t', elem).text();
    console.log(a);
//      console.log(pd.xml($(elem, {xmlMode:true}).xml()));
  });

  var str2 = $slide1.xml();
  zip.file('ppt/slides/slide1.xml', str2);

  var buffer = zip.generate({type:"nodebuffer"});

  fs.writeFile(OUTFILE, buffer, function(err) {
    if (err) throw err;
  });


});
