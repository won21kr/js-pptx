var fs = require('fs');
var xml2js = require('xml2js')
var arguments = process.argv.slice(2);
var INFILE = arguments[0];
var OUTFILE = arguments[1];

fs.readFile(INFILE, 'utf8', function(err, xml) {
  if (err) throw(err);
  xml2js.parseString(xml, function(err, js) {
    if (err) throw(err);
    if (OUTFILE) {
      var txt = "module.exports = " + JSON.stringify(js,null,4);
      fs.writeFile(OUTFILE, txt, 'utf8', function(err) {
        if (err) throw(err);
        console.log("File written to "+OUTFILE);
      } )
    }
    else console.log(JSON.stringify(js,null,4));
  })

});
