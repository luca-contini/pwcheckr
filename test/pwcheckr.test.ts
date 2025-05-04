import { test } from 'node:test';
import assert from 'node:assert/strict';

import {
    isDoc,
    isDocPasswordProtected,
    isDocx,
    isDocxPasswordProtected,
    isPDF,
    isPDFPasswordProtected,
    isPpt,
    isPptPasswordProtected,
    isPptx, isPptxPasswordProtected,
    isXls,
    isXlsPasswordProtected,
    isXlsx,
    isXlsxPasswordProtected,
} from '../src/pwcheckr';

async function fixture(filename: string): Promise<File> {
    const fs = await import('node:fs/promises');
    const path = await import('node:path');

    const filePath = path.resolve('fixtures', filename);
    const buffer = await fs.readFile(filePath);

    return new File([buffer], filename);
}

test('should detect .docx files', async () => {
    const file = await fixture('google-unprotected.docx');

    assert.equal(isDoc(file), false);
    assert.equal(isDocx(file), true);
    assert.equal(await isDocxPasswordProtected(file), false);

    /*
     * Note about `file.type`: based on the current implementation, browsers won't actually read the bytestream of a file to determine its media type.
     * It is assumed based on the file extension; a PNG image file renamed to .txt would give "text/plain" and not "image/png". Moreover, blob.type is
     * generally reliable only for common file types like images, HTML documents, audio and video. Uncommon file extensions would return an empty string.
     * Client configuration (for instance, the Windows Registry) may result in unexpected values even for common types.
     *
     * Developers are advised not to rely on this property as a sole validation scheme.
     */
    assert.equal(isDocx(file, true), false); // Assert false because of the above.
});

test('should detect password-protected .docx (CDFV2 / Office 2007+)', async () => {
    const file = await fixture('google-protected.docx');

    assert.equal(isDoc(file), false);

    assert.equal(isDocx(file), true);
    assert.equal(await isDocxPasswordProtected(file), true);
    assert.equal(await isDocxPasswordProtected(file, true), false);
});

test('should handle empty .doc files', async () => {
    const file = await fixture('empty.doc');

    assert.equal(isDocx(file), false);

    assert.equal(isDoc(file), true);
    assert.equal(isDoc(file, true), false);
    assert.equal(await isDocPasswordProtected(file), false);
});

test('should handle empty .docx files', async () => {
    const file = await fixture('empty.docx');

    assert.equal(isDoc(file), false);

    assert.equal(isDocx(file), true);
    assert.equal(isDocx(file, true), false);
    assert.equal(await isDocxPasswordProtected(file), false);
});

test('should handle empty .pdf files', async () => {
    const file = await fixture('empty.pdf');

    assert.equal(isPDF(file), true);
    assert.equal(isPDF(file, true), false);
    assert.equal(await isPDFPasswordProtected(file), false);
});

test('should handle empty .ppt files', async () => {
    const file = await fixture('empty.ppt');

    assert.equal(isPptx(file), false);

    assert.equal(isPpt(file), true);
    assert.equal(isPpt(file, true), false);
    assert.equal(await isPptPasswordProtected(file), false);
});

test('should handle empty .pptx files', async () => {
    const file = await fixture('empty.pptx');

    assert.equal(isPpt(file), false);

    assert.equal(isPptx(file), true);
    assert.equal(isPptx(file, true), false);
    assert.equal(await isPptxPasswordProtected(file), false);
});

test('should handle empty .xls files', async () => {
    const file = await fixture('empty.xls');

    assert.equal(isXlsx(file), false);

    assert.equal(isXls(file), true);
    assert.equal(isXls(file, true), false);
    assert.equal(await isXlsPasswordProtected(file), false);
});

test('should handle empty .xlsx files', async () => {
    const file = await fixture('empty.xlsx');

    assert.equal(isXls(file), false);

    assert.equal(isXlsx(file), true);
    assert.equal(isXlsx(file, true), false);
    assert.equal(await isXlsxPasswordProtected(file), false);
});