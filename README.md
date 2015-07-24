# js-pptx
Pure Javascript reader/writer for PowerPoint, for use in Node.js or the browser.

# Design goals
* Read/edit/author PowerPoint .pptx files
* Pure Javascript
* Runs in browser and/or Node.js
* Lightweight
* Friendly API for basic tasks, like text, shapes, charts, tables
* Access to raw XML for when you need to be very specific
* Rigorous test suite

# Inspiration / Motivation
Inspired by [officegen](https://github.com/ZivBarber/officegen),
which creates pptx with text/shapes/images/tables/charts wonderfully (but does not read existing PPT files).

Also inspired by [js-xlsx](https://github.com/SheetJS/js-xlsx)
which reads/writes XLSX/XLS/XLSB, works in the browser and Node.js, and has an incredibly
thorough test suite (but does not read or write PowerPoint).

Motivated by desire to read and modify existing presentations, to inherit their themes, layouts and possibly content,
and work in the browser if possible.

# Design Philosophy

The design approach here uses:
* `JSZip` to unzip an existing `.pptx` file and zip it back,
* `xml2js` to convert the XML to Javascript and back to XML.

This library currently assumes it's starting from an existing presentation, and doesn't create one from scratch.
This allows you to use existing themes, styles and layouts.

Converting to Javascript allows the content to be manipulated programmatically.  For each major entity, a Javascript class is created,
such as:
 * PPTX.Presentation
 * PPTX.Slide
 * PPTX.Shape
 * PPTX.spPr  // ShapeProperties
 * etc.

These classes allow properties to be set, and chained in a manner similar to d3 or jQuery.

Right now, this uses English names for high-level constructs (e.g. `Presentation` and `Slide`),
but uses names that directly mirror the OpenXML tagNames for lower level constructs (e.g.  `spPr` for ShapeProperties).

The challenge is it'll be a lot easier to extend the library if we follow the OpenXML tag names,
but the OpenXML tag names are so cryptic that they don't make great names for a Javascript library.

So we default to using the English name is used when returning objects even if the object has a cryptic class name, e.g.:
* `Slide.getShapes()` returns an array of `Shape` objects and
* `Shape.shapeProperties()` returns an `spPr` object.

Ideally would be consistent, and am working out which way to go.  Advice is welcome!



# Install

In node.js
```
npm install protobi/js-pptx
```

In the browser:
```
<script src="js-pptx.js"></script>  // will use browserify but right now not yet implemented
```

# Dependencies
* [xml2js](https://github.com/nfarina/xmldoc)
* [async](https://github.com/caolan/async)
* [jszip](https://stuk.github.io/jszip)

# Usage

```js
var PPTX = require('../lib/pptx');
var fs = require('fs');

var INFILE = './test/files/parts3.pptx';
var OUTFILE = './test/files/parts3-a.pptx';

fs.readFile(INFILE, function (err, data) {
  if (err) throw err;

  var pptx = new PPTX.Presentation();
  pptx.load(data, function (err) {

    var slide1 = pptx.getSlide('slide1');
    var shapes = slide1.getShapes();

    var shapes = slide1.getShapes()
    shapes[3]
        .text("Now it's a trapezoid")
        .shapeProperties()
        .x(PPTX.emu.inch(1))
        .y(PPTX.emu.inch(1))
        .cx(PPTX.emu.inch(2))
        .cy(PPTX.emu.inch(0.75))fs.writeFile(OUTFILE, pptx.toBuffer(), function (err) {
        .prstGeom('trapezoid');  if (err) throw err;
    });
  });
});

```

# Contribute

###Test:
`npm run test`

###Build:
`npm run build`

###Minify:
`npm run minify`

###All:
`npm run all`



# Current status:
Embryonic - not at all ready for use.




