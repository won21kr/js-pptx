var JSZip = require('jszip');
var xmldoc = require('xmldoc');

var fs = require('fs');

var zip = new JSZip();
var content = fs.readFileSync(__dirname + '/lab/ex1/ex1.pptx');

zip.load(content);

var xml_slide1 = zip.file("ppt/slides/slide2.xml").asText();
var doc_slide1 = new xmldoc.XmlDocument(xml_slide1);

doc_slide1
  .childNamed('p:cSld')
  .childNamed('p:spTree')
  .childrenNamed('p:sp')[0]
  .childNamed('p:txBody')
  .childNamed('a:p')
  .childNamed('a:r')
  .childNamed('a:t').val  = "Gotcha!"

zip.file("ppt/slides/slide2.xml", doc_slide1.toString())


var buffer = zip.generate({type:"nodebuffer"});
fs.writeFile("/tmp/test.pptx", buffer, function(err) {
  if (err) throw err;
  console.log("Done")
});
//.map(function(child){ return child.name}).children
var str  = '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0" smtClean="0"/><a:t>Chart 1 Body</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>';