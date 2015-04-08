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

# Inspired by
Inspired by [officegen](https://github.com/ZivBarber/officegen), which creates pptx with text/shapes/images/tables wonderfully, but does not work in the browser
or read/edit existing PPT files.

Also inspired by [js-xlsx](https://github.com/SheetJS/js-xlsx) which reads/writes XLSX/XLS/XLSB, works in the browser and Node.js, and has an incredibly
thorough test suite, but does not read/write PowerPoint files.

# Install

In node.js
`npm install protobi/jszip`
`npm install protobi/xmldoc`
`npm install protobi/js-pptx`

In the browser:

`<script src="xmldoc.min.js"></script>`
`<script src="xlsx.min.js"></script>`
`<script src="js-pptx.js"></script>`


# Dependencies
* [xmldoc](https://github.com/nfarina/xmldoc)
* [jszip](https://stuk.github.io/jszip)

# Contribute

Test:
`npm run test`

Build:
`npm run build`

Minify:
`npm run minify`

All:
`npm run all`



# Current status:
Embryonic - not at all ready for use.




