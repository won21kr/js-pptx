"use strict";

var JSZip = require('jszip');
var async = require('async');
var xml2js = require('xml2js');

var PPTX = function(object) {
  this.content = {};
};

// fundamentally asynchronous because xml2js.parseString() is async
PPTX.prototype.load = function (data, done) {
  var zip = new JSZip(data);

  var content = this.content;
  async.each(Object.keys(zip.files), function (key, callback) {
    var ext = key.substr(key.indexOf('.'));
    if (ext == '.xml' || ext == '.xml.rels') {
      var xml = zip.file(key).asText();
      xml2js.parseString(xml, function (err, js) {
        pptx[key] = js;
        callback(null);
      });
    }
    else {
      pptx[key] = zip.file(key).asText();
      callback(null);
    }
  }, done);
};

PPTX.prototype.toJSON = function() {
  return this.content;
};


PPTX.prototype.toBuffer = function () {
  var zip2 = new JSZip();
  var content = this.content;
  for (var key in content) {
    if (content.hasOwnProperty(key)) {
      var ext = key.substr(key.indexOf('.'));
      if (ext == '.xml' || ext == '.xml.rels') {
        var builder = new xml2js.Builder();
        var xml2 = (builder.buildObject(content[key]));
        zip2.file(key, xml2);
      }
      else {
        zip2.file(key, content[key]);
      }
    }
  }
  var buffer = zip2.generate({type: "nodebuffer"});
  return buffer;
};

module.exports = PPTX;


