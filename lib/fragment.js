var fs = require('fs');
var xml2js = require('xml2js');

module.exports = {
  fromXml: function(name, callback) {
    fs.readFile(__dirname + '/fragments/' + name, 'utf8', function(err, xml) {
      if (err) callback(err);
      xml2js.parseString(xml, function (err, js) {
        callback(null, js);
      });
    })
  },

  fromBinary: function(name, callback) {
    fs.readFile(__dirname + '/fragments/' + name, function(err, data) {
      if (err) callback(err);
      console.log("Data file read: "+data.length   + "  typeof "+(data instanceof Buffer));
      callback(null, data );
    })
  },

  get_workbook : function(name, callback) {
    fs.readFile(__dirname + '/fragments/Microsoft_Excel_Sheet1.xlsx', callback);
  }
}