var PPTX = require('./pptx');
var xmldoc = require('xmldoc');
var fs = require('fs');

var pptx = new PPTX();
pptx.readFile(__dirname + '/lab/ex1/ex1.pptx', function(err) {

  var doc_slide1 = pptx.getSlideDoc(2);
  //var doc_slide1 = new xmldoc.XmlDocument(xml_slide1);

  // abstract this into a Slide and Shape class
  // see http://www.officeopenxml.com/drwSp-text.php
  doc_slide1
    .childNamed('p:cSld') // slide
    .childNamed('p:spTree') // shapeTree
    .childrenNamed('p:sp')[0] // shapes
    .childNamed('p:txBody') // text contained within shape
    .childNamed('a:p')
    .childNamed('a:r')
    .childNamed('a:t').val  = "Hi!";

  pptx.writeFile("/tmp/test.pptx", function(err) {
    if (err) throw err;
  });

});

var str  = '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>Chart 1 Body</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>';



