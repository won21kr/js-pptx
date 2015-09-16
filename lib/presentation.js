var JSZip = require('jszip'); // this works on browser
var async = require('async'); // this works on browser
var xml2js = require('xml2js'); // this works on browser?
var XmlNode = require('./xmlnode');

var Slide = require('./slide');

var Presentation = function (object) {
  this.content = {};
};

// fundamentally asynchronous because xml2js.parseString() is async
Presentation.prototype.load = function (data, done) {
  var zip = new JSZip(data);

  var content = this.content;
  async.each(Object.keys(zip.files), function (key, callback) {
    var ext = key.substr(key.lastIndexOf('.'));
    if (ext == '.xml' || ext == '.rels') {
      var xml = zip.file(key).asText();

      xml2js.parseString(xml, function (err, js) {
        content[key] = js;
        callback(null);
      });
    }
    else {
      content[key] = zip.file(key).asText();
      callback(null);
    }
  }, done);
};

Presentation.prototype.toJSON = function () {
  return this.content;
};


Presentation.prototype.toBuffer = function () {
  var zip2 = new JSZip();
  var content = this.content;
  for (var key in content) {
    if (content.hasOwnProperty(key)) {
      var ext = key.substr(key.lastIndexOf('.'));
      if (ext == '.xml' || ext == '.rels') {
        var builder = new xml2js.Builder({renderOpts: {pretty: false}});
        var xml2 = (builder.buildObject(content[key]));
        zip2.file(key, xml2);
      }
      else {
        zip2.file(key, content[key]);
      }
    }
  }
//  zip2.file("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>1</TotalTime><Words>72</Words><Application>Microsoft Macintosh PowerPoint</Application><PresentationFormat>On-screen Show (4:3)</PresentationFormat><Paragraphs>12</Paragraphs><Slides>3</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>3</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="4" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr><vt:lpstr>This is the title</vt:lpstr><vt:lpstr>This is the title</vt:lpstr><vt:lpstr>This is the title</vt:lpstr></vt:vector></TitlesOfParts><Company>Proven, Inc.</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>');
  zip2.file("docProps/app.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><TotalTime>0</TotalTime><Words>0</Words><Application>Microsoft Macintosh PowerPoint</Application><PresentationFormat>On-screen Show (4:3)</PresentationFormat><Paragraphs>0</Paragraphs><Slides>2</Slides><Notes>0</Notes><HiddenSlides>0</HiddenSlides><MMClips>0</MMClips><ScaleCrop>false</ScaleCrop><HeadingPairs><vt:vector size="4" baseType="variant"><vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant><vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant><vt:variant><vt:i4>2</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="3" baseType="lpstr"><vt:lpstr>Office Theme</vt:lpstr><vt:lpstr>PowerPoint Presentation</vt:lpstr><vt:lpstr>PowerPoint Presentation</vt:lpstr></vt:vector></TitlesOfParts><Company>Proven, Inc.</Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>14.0000</AppVersion></Properties>')
  var buffer = zip2.generate({type: "nodebuffer"});
  return buffer;
};

Presentation.prototype.registerChart = function(chartName, content) {
  this.content['ppt/charts/' + chartName + '.xml'] = content;

  // '[Content_Types].xml' .. add references to the chart and the spreadsheet
  this.content["[Content_Types].xml"]["Types"]["Override"].push(XmlNode()
      .attr('PartName', "/ppt/charts/" + chartName + ".xml")
      .attr('ContentType', "application/vnd.openxmlformats-officedocument.drawingml.chart+xml")
      .el
  );

  var defaults = this.content["[Content_Types].xml"]["Types"]["Default"].filter(function (el) {
    return el['$']['Extension'] == 'xlsx'
  });

  if (defaults.length == 0) {
    this.content["[Content_Types].xml"]["Types"]["Default"].push(XmlNode()
        .attr('Extension', 'xlsx')
        .attr('ContentType', "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        .el
    );
  }
}

Presentation.prototype.registerChartWorkbook = function(chartName, workbookContent) {

  var numWorksheets = this.getWorksheetCount();
  var worksheetName = 'Microsoft_Excel_Sheet' + (numWorksheets + 1) + '.xlsx';


  this.content["ppt/embeddings/" + worksheetName] = workbookContent;

  // ppt/charts/_rels/chart1.xml.rels
  this.content["ppt/charts/_rels/" + chartName + ".xml.rels"] = XmlNode().setChild("Relationships", XmlNode()
      .attr({
        'xmlns': "http://schemas.openxmlformats.org/package/2006/relationships"
      })
      .addChild('Relationship', XmlNode().attr({
        "Id": "rId1",
        "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package",
        "Target": "../embeddings/" + worksheetName
      }))
  ).el;
}

Presentation.prototype.getSlideCount = function () {
  return Object.keys(this.content).filter(function (key) {
    return key.substr(0, 16) === "ppt/slides/slide"
  }).length;
}

Presentation.prototype.getChartCount = function () {
  return Object.keys(this.content).filter(function (key) {
    return key.substr(0, 16) === "ppt/charts/chart"
  }).length;
}

Presentation.prototype.getWorksheetCount = function () {
  return Object.keys(this.content).filter(function (key) {
    return key.substr(0, 36) === "ppt/embeddings/Microsoft_Excel_Sheet"
  }).length;
}



Presentation.prototype.getSlide = function (name) {
  return new Slide({content: this.content['ppt/slides/' + name + '.xml'], presentation: this, name: name});
}

Presentation.prototype.addSlide = function (layoutName) {
  var slideName = "slide" + (this.getSlideCount() + 1);

  var layoutKey = "ppt/slideLayouts/" + layoutName + ".xml";
  var slideKey = "ppt/slides/" + slideName + ".xml";
  var relsKey = "ppt/slides/_rels/" + slideName + ".xml.rels";


  // create slide
  //  var slideContent = this.content[layoutKey]["p:sldLayout"];



  //var sld = this.content["ppt/slides/slide1.xml"];   // this is cheating, copying an existing slide
  var sld =  this.content[layoutKey]["p:sldLayout"];
  delete sld['$']["preserve"];
  delete sld['$']["type"];

  var slideContent = {
    "p:sld" : sld
  };


  slideContent = JSON.parse(JSON.stringify(slideContent));

  this.content[slideKey] = slideContent; //{ "p:sld": slideContent};


  var slide = new Slide({content: slideContent, presentation: this, name: slideName});

  // add to [Content Types].xml
  this.content["[Content_Types].xml"]["Types"]["Override"].push({
    "$": {
      "PartName": "/ppt/slides/" + slideName + ".xml",
      "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml"
    }
  });

  //add it rels to slidelayout
  this.content[relsKey] = {
    "Relationships": {
      "$": {
        "xmlns": "http://schemas.openxmlformats.org/package/2006/relationships"
      },
      "Relationship": [
        {
          "$": {
            "Id": "rId1",
            "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
            "Target": "../slideLayouts/" + layoutName + ".xml"
          }
        }
      ]
    }
  };

  // add it to ppt/_rels/presentation.xml.rels
  var rId = "rId" + (this.content["ppt/_rels/presentation.xml.rels"]["Relationships"]["Relationship"].length + 1);
  this.content["ppt/_rels/presentation.xml.rels"]["Relationships"]["Relationship"].push({
    "$": {
      "Id": rId,
      "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
      "Target": "slides/" + slideName + ".xml"
    }
  });

  // add it to ppt/presentation.xml
  var maxId = 0;
  this.content["ppt/presentation.xml"]["p:presentation"]["p:sldIdLst"][0]["p:sldId"].forEach(function (ob) {
    if (+ob["$"]["id"] > maxId) maxId = +ob["$"]["id"]
  })
  this.content["ppt/presentation.xml"]["p:presentation"]["p:sldIdLst"][0]["p:sldId"].push({
    "$": {
      "id": "" + (+maxId + 1),
      "r:id": rId
    }
  });

  // increment slidecount
  var sldCount = this.getSlideCount();
  this.content["docProps/app.xml"]["Properties"]["Slides"][0] = sldCount ;


  return slide;
}

module.exports = Presentation;