'use strict';

const LinkManager = require('./linkManager');
const wrapper = require('docxtemplater/js/module-wrapper');

const moduleName = 'linkmodule';

class LinkModule {
  constructor() {
    this.name = 'LinkModule';
    this.prefix = '$';
  }

  optionsTransformer(options, docxtemplater) {
    const relsFiles = docxtemplater.zip
      .file(/\.xml\.rels/)
      .concat(docxtemplater.zip.file(/\[Content_Types\].xml/))
      .map((file) => file.name);
    this.fileTypeConfig = docxtemplater.fileTypeConfig;
    this.fileType = docxtemplater.fileType;
    this.zip = docxtemplater.zip;
    options.xmlFileNames = options.xmlFileNames.concat(relsFiles);
    return options;
  }

  set(obj) {
    if (obj.zip) {
      this.zip = obj.zip;
    }
    if (obj.compiled) {
      this.compiled = obj.compiled;
    }
    if (obj.xmlDocuments) {
      this.xmlDocuments = obj.xmlDocuments;
    }
    if (obj.data != null) {
      this.data = obj.data;
    }
  }

  matchers() {
    return [[this.prefix, moduleName]];
  }

  getRenderedMap(mapper) {
    return Object.keys(this.compiled).reduce((mapper, from) => {
      mapper[from] = { from, data: this.data };
      return mapper;
    }, mapper);
  }

  render(part, { scopeManager, filePath }) {
    if (part.module !== moduleName) {
      return null;
    }

    const linkManager = new LinkManager(this.zip, filePath, this.xmlDocuments, this.fileType);
    const tagValue = scopeManager.getValue(part.value, { part });
    const rId = linkManager.addLinkRels(tagValue);

    if (!tagValue) {
      return { value: this.fileTypeConfig.tagTextXml };
    }

    const value = `</w:t></w:r><w:hyperlink r:id="${rId}" w:history="1"><w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/><w:r w:rsidR="00052F25" w:rsidRPr="00052F25"><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>${tagValue}</w:t></w:r></w:hyperlink><w:r><w:t xml:space="preserve">`;
    return { value };
  }
}

module.exports = () => wrapper(new LinkModule());
