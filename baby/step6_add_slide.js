"use strict";


var JSZip = require('jszip');
var fs = require("fs");
var cheerio = require('cheerio');
var xmldoc = require('xmldoc');
var pd = require('pretty-data').pd;
var pretty = function (xml) {
  return pd.xml(xml);
}

var INFILE = './lab/parts/parts.pptx';
var OUTFILE = './lab/parts2/parts2.pptx';
var REFERENCE = './lab/parts3/parts3.pptx';


function $(xml) {
  return cheerio.load(xml, {xmlMode: true})
}

function Slide(xml) {
  this.$el = cheerio.load(xml, {xmlMode: true});
  return this;
};
Slide.prototype.xml = function () {
  return this.$el.xml();
}

var zip3 = new JSZip(fs.readFileSync(REFERENCE))

fs.readFile(INFILE, function (err, data) {
  if (err) throw err;
  var zip = new JSZip(data);



  var slide1xml = zip.file('ppt/slides/slide1.xml').asText();
  var $slide1 = cheerio.load(slide1xml, {xmlMode: true})


  // 1.  add the slide to ppt/slides/slideN.xml
  zip.file('ppt/slides/slide2.xml', slide1xml);


  // 2.  add entry to ppt/_rels/presentation.xml.rels

  var $ppt_rels = $(zip.file('ppt/_rels/presentation.xml.rels').asText());
  var rId3 = '<Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>';
  $ppt_rels('Relationships').append(rId3)
  zip.file('ppt/_rels/presentation.xml.rels', $ppt_rels.xml());


  // 3.  add reference to ppt/slides/_rels/
  var xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
  xml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/></Relationships>';
  zip.file('ppt/slides/_rels/slide2.xml.rels', xml)

  // 4.  add slide to ppt/presentation.xml
  var pres = zip.file('ppt/presentation.xml').asText();
  var $pres = $(pres);
  $pres('p\\:sldIdLst').append('<p:sldId id="257" r:id="rId3"/>')
  zip.file('ppt/presentation.xml', $pres.xml())


  // HACK!   set docProps

  zip.file('docProps/app.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>1</TotalTime><Words>48</Words><Application>Microsoft Macintosh PowerPoint</Application><PresentationFormat>On-screen Show (4:3)</PresentationFormat><Paragraphs>8</Paragraphs><Slides>2</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr><vt:lpstr>This is the title</vt:lpstr><vt:lpstr>This is the title</vt:lpstr></vt:vector></TitlesOfParts><Company>Proven, Inc.</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>')
  zip.file('ppt/viewProps.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" lastView="sldThumbnailView"><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1" showGuides="1"><p:cViewPr varScale="1"><p:scale><a:sx n="95" d="100"/><a:sy n="95" d="100"/></p:scale><p:origin x="-1384" y="-104"/></p:cViewPr><p:guideLst><p:guide orient="horz" pos="2160"/><p:guide pos="2880"/></p:guideLst></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="100" d="100"/><a:sy n="100" d="100"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>')



  var $ct = $(zip.file('[Content_Types].xml').asText());
  $ct('Types').append('<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>');
//  console.log(pretty($ct.xml()));
  zip.file('[Content_Types].xml', $ct.xml())
  var $slidIdList =
  console.log($($pres('p\\:sldIdLst').children()[0]).xml())



    // SYSTEMATICALLY FIND PROBLEM
  var diffs = [
//    '[Content_Types].xml',
//    'ppt/_rels/presentation.xml.rels',
//    'docProps/core.xml',
//    'ppt/slides/slide2.xml'
  ];
//  var diffs = Object.keys(zip3.files);
//  console.log(diffs);
  diffs.forEach(function(key) {
    zip.file(key, zip3.file(key).asArrayBuffer())
  })
//

  var buffer = zip.generate({type: "nodebuffer"});

  fs.writeFile(OUTFILE, buffer, function (err) {
    if (err) throw err;
  });
//

});
