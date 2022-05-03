# Open Source docxtemplater link module

## Installation

You first need to install docxtemplater by following its [installation guide](https://docxtemplater.readthedocs.io/en/latest/installation.html).

For Node.js install this package:

```bash
npm install docxtemplater-link-module-free
```

For the browser find builds in `build/` directory.

Alternatively, you can create your own build from the sources:

```bash
npm run compile
npm run browserify
npm run uglify
```

## Usage

Assuming your **docx** template contains only the text `{^link}`:

```javascript
var LinkModule = require("open-docxtemplater-link-module");
var linkModule = new LinkModule();

var zip = new JSZip(content);
var doc = new Docxtemplater()
  .attachModule(linkModule)
  .loadZip(zip)
  .setData({ link: "https://google.fr" })
  .render();

var buffer = doc.getZip().generate({ type: "nodebuffer" });

fs.writeFile("test.docx", buffer);
```
