  "use strict";


  var JSZip = require('jszip');
  var fs = require("fs");

  //var FILE1 = './test/files/parts3-b.pptx';
  //var FILE2 = './test/files/parts3-a.pptx';

  var FILE1 = './lab/p0/protobi-2015-09-03 17.06.pptx';
  var FILE2 = './lab/p1/protobi-2015-09-03 17.06.pptx';

  var zip1 = new JSZip(fs.readFileSync(FILE1));
  var zip2 = new JSZip(fs.readFileSync(FILE2));
  Object.keys(zip1.files).forEach(function (key) {
    var str1 = zip1.file(key).asText().replace(/\n|\s/ig, '');
    var str2 = zip2.file(key) ? zip2.file(key).asText().replace(/\n|\s/ig, '') : "";

  //  if (str1 != str2) {
    if (str1.length != str2.length) {

  //  if (str1.length != str2.length) {
      console.log(key);
      console.log(str1.length, str2.length, str1.length == str2.length)
    }

  });
