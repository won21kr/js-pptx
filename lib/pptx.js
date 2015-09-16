"use strict";

console.json = function(obj) { console.log(JSON.stringify(obj, null,4)); }
module.exports = {
  Presentation: require('./presentation'),
  Slide: require('./slide'),
  Shape: require('./shape'),
//  spPr: require('./spPr'),
  emu: {
    inch: function(val) { return Math.floor(val * 914400); },
    point: function(val) { return Math.floor(val * 914400 / 72); },
    px: function(val) { return Math.floor(val * 914400 / 72); },
    cm: function(val) { return Math.floor(val * 360000); }
  }
};




