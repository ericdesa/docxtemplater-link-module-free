'use strict';

const DocUtils = require('docxtemplater').DocUtils;

const rels = {
  getPrefix(fileType) {
    return fileType === 'docx' ? 'word' : 'ppt';
  },
  getFileTypeName(fileType) {
    return fileType === 'docx' ? 'document' : 'presentation';
  },
  getRelsFileName(fileName) {
    return fileName.replace(/^.*?([a-zA-Z0-9]+)\.xml$/, '$1') + '.xml.rels';
  },
  getRelsFilePath(fileName, fileType) {
    const relsFileName = rels.getRelsFileName(fileName);
    const prefix = fileType === 'pptx' ? 'ppt/slides' : 'word';
    return `${prefix}/_rels/${relsFileName}`;
  },
};

module.exports = class LinkManager {
  constructor(zip, fileName, xmlDocuments, fileType) {
    this.nbLinks = 1;
    this.fileName = fileName;
    this.prefix = rels.getPrefix(fileType);
    this.zip = zip;
    this.xmlDocuments = xmlDocuments;
    this.fileTypeName = rels.getFileTypeName(fileType);
    const relsFilePath = rels.getRelsFilePath(fileName, fileType);
    this.relsDoc = xmlDocuments[relsFilePath] || this.createEmptyRelsDoc(xmlDocuments, relsFilePath);
  }

  createEmptyRelsDoc(xmlDocuments, relsFileName) {
    const mainRels = this.prefix + '/_rels/' + this.fileTypeName + '.xml.rels';
    const doc = xmlDocuments[mainRels];
    if (!doc) {
      const err = new Error('Could not copy from empty relsdoc');
      err.properties = {
        mainRels,
        relsFileName,
        files: Object.keys(this.zip.files),
      };
      throw err;
    }
    const relsDoc = DocUtils.str2xml(DocUtils.xml2str(doc));
    const relationships = relsDoc.getElementsByTagName('Relationships')[0];
    const relationshipChilds = relationships.getElementsByTagName('Relationship');
    for (let i = 0, l = relationshipChilds.length; i < l; i++) {
      relationships.removeChild(relationshipChilds[i]);
    }
    xmlDocuments[relsFileName] = relsDoc;
    return relsDoc;
  }

  loadLinkRels() {
    const iterable = this.relsDoc.getElementsByTagName('Relationship');
    return Array.prototype.reduce.call(
      iterable,
      function (max, relationship) {
        const id = relationship.getAttribute('Id');
        if (/^rId[0-9]+$/.test(id)) {
          return Math.max(max, parseInt(id.substr(3), 10));
        }
        return max;
      },
      0,
    );
  }

  addLinkRels(linkValue) {
    const relationships = this.relsDoc.getElementsByTagName('Relationships')[0];
    const relationshipChilds = relationships.getElementsByTagName('Relationship');

    const newTag = this.relsDoc.createElement('Relationship');
    newTag.namespaceURI = null;
    const rId = `link_generated_${relationshipChilds.length + 1}`;
    newTag.setAttribute('Id', rId);
    newTag.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink');
    newTag.setAttribute('Target', linkValue);
    newTag.setAttribute('TargetMode', 'External');
    relationships.appendChild(newTag);
    return rId;
  }
};
