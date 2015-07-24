
var JSZip = require('jszip'); // this works on browser
var async = require('async'); // this works on browser
var xml2js = require('xml2js'); // this works on browser?


var Slide = require('./slide');

var Presentation = function (object) {
  this.content = {};
};

// fundamentally asynchronous because xml2js.parseString() is async
Presentation.prototype.load = function (data, done) {
  var zip = new JSZip(data);

  var content = this.content;
  async.each(Object.keys(zip.files), function (key, callback) {
    var ext = key.substr(key.indexOf('.'));
    if (ext == '.xml' || ext == '.xml.rels') {
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
      var ext = key.substr(key.indexOf('.'));
      if (ext == '.xml' || ext == '.xml.rels') {
        var builder = new xml2js.Builder({renderOpts: {pretty: false}});
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

Presentation.prototype.getSlide = function (name) {
  return new Slide(this.content['ppt/slides/' + name + '.xml']);
}


module.exports = Presentation;