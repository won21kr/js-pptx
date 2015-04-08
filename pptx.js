// TODO
// Test suite
// Wrap PPT into a class object
// Abstract for Node and Browser
// Create utility methods for text, title, properties, shapes
// Embed charts
// Browserify


var JSZip = require('jszip');
var xmldoc = require('xmldoc');

var fs = require('fs');



function PPTX() {
  var zip = new JSZip();
  var _docs = Object.create(null);

  return {
    read: function(content) {
      zip.load(content);
      return this;
    },

    readFile: function(filename, callback) {
      var self = this;
      fs.readFile(filename, function(err, content) {
        if (err) { return callback(err); }
        try {
          self.read(content);
        }
        catch (e) { return callback(e); }
        callback();

      });
      return this;
    },

    finalize: function() {
      for (key in _docs) {
        zip.file(key, _docs[key].toString())
      }
    },

    write: function() {
      this.finalize();
      return zip.generate({type:"nodebuffer"});
    },

    writeFile: function(filename, callback) {
      var buffer = this.write();
      fs.writeFile(filename, buffer, callback);
    },

    // returns XML document for slide at specified index, starting with 1
    getSlideDoc: function(slideIndex) {

      var key = 'ppt/slides/slide' + slideIndex + '.xml';
      var doc = _docs[key];
      if (!doc) {
        var xml = zip.file(key).asText();
        doc = _docs[key] = new xmldoc.XmlDocument(xml);
      }
      return doc;
    }


  }
}

module.exports = PPTX;