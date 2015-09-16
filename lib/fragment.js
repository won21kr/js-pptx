var fs = require('fs');
var xml2js = require('xml2js');

module.exports = {
  fromXml: function(name, callback) {
    fs.readFile(__dirname + '/fragments/' + name, 'utf8', function(err, xml) {
      if (err) callback(err);
      xml2js.parseString(xml,{explicitArray : false}, function (err, js) {
        callback(null, js);
      });
    })
  },

  fromBinary: function(name, callback) {
    fs.readFile(__dirname + '/fragments/' + name, function(err, data) {
      if (err) callback(err);
      callback(null, data );
    })
  }
}