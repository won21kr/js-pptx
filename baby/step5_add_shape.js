"use strict";


var JSZip = require('jszip');
var fs = require("fs");
var cheerio = require('cheerio');
var xmldoc = require('xmldoc');
var pd = require('pretty-data').pd;
var pretty = function(xml) { return pd.xml(xml); }

var INFILE = './lab/parts/parts.pptx';
var OUTFILE = './lab/parts/parts2.pptx';

function $(xml) { return cheerio.load(xml, {xmlMode: true})}

function Slide(xml) {
  this.$el = cheerio.load(xml, {xmlMode: true});
  return this;
};
Slide.prototype.xml = function() { return this.$el.xml(); }

fs.readFile(INFILE, function(err, data) {
  if (err) throw err;
  var zip = new JSZip(data);

  var slide1xml = zip.file('ppt/slides/slide1.xml').asText();
  var $slide1 = cheerio.load(slide1xml, {xmlMode: true})

  var $shapes = $slide1('p\\:sp');

//  textBoxes.each(function(i, elem) {
//    console.log(i + "-----------------------------")
//    console.log(cheerio.load(elem).xml())
//  })
//
  var $oval =$shapes[3];
//  console.log(($($oval)));

  var str = '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Oval 5"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="6578600" y="787400"/><a:ext cx="1181100" cy="1181100"/></a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr><p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="3"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="2"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>Another circle</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>';

  $slide1('p\\:spTree').append(str);


  var str2 = $slide1.xml();
  zip.file('ppt/slides/slide1.xml', str2);

  var buffer = zip.generate({type:"nodebuffer"});

  fs.writeFile(OUTFILE, buffer, function(err) {
    if (err) throw err;
  });




//
//  var $ovalText = $($oval)('a\\:t');
//
//
//  console.log($ovalText.text());
//  console.log($($ovalText[0]).xml() );

//  console.log(($($oval).xml()));


});
