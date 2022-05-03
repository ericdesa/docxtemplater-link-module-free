'use strict';
/* eslint-disable no-console */

const Docxtemplater = require('docxtemplater');
const path = require('path');
const PizZip = require('PizZip');
const LinkModule = require('./index.js');
const testutils = require('docxtemplater/js/tests/utils');
const shouldBeSame = testutils.shouldBeSame;

const fileNames = ['test.docx'];

beforeEach(function () {
	this.opts = {};

	this.loadAndRender = function () {
		// const file = testutils.createDoc(this.name);
		// const zip = new PizZip(file.loadedContent);
		// const docxTemplate = new Docxtemplater(zip, {
		// 	module: [new LinkModule()],
		// });
		// docxTemplate.setData(data);
		// docxTemplate.render();
		// const buffer = docxTemplate.getZip().generate({
		// 	type: 'nodebuffer',
		// 	mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
		// });
		// fs.writeFileSync('../examples/' + this.expectedName, buffer);
		// shouldBeSame({ buffer, expectedName: this.expectedName });
	};
});

function testStart() {
	describe('{^link}', function () {
		it('should interpolate link', function () {
			this.name = 'test.docx';
			this.expectedName = 'expectedTest.docx';
			this.data = {};
			this.loadAndRender();
		});
	});
}

testutils.setExamplesDirectory(path.resolve(__dirname, '..', 'examples'));
testutils.setStartFunction(testStart);
fileNames.forEach(function (filename) {
	console.log(path.resolve(__dirname, '..', 'examples'), filename);
	testutils.loadFile(filename, testutils.loadDocument);
});
// testutils.start();
