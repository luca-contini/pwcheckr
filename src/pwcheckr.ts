import { parse as parseCFB } from 'cfb';

/**
 * Return whether a file does not have a true OOXML container (ZIP-based).
 *
 * Since the source may be detected as OLE2/Compound File Binary Format (like .doc, but renamed with a .docx extension),
 * this performs a stricter check to ensure a file is zipped.
 */
async function isOOXML(file: File): Promise<boolean> {
    const header = await readFileHeader(file, 4);

    // ZIP magic numbers: 0x50 0x4B 0x03 0x04
    return header[0] === 0x50 && header[1] === 0x4B && header[2] === 0x03 && header[3] === 0x04;
}

/**
 * Return whether the provided file contains encryption information (binary CFBF format).
 *
 * This is a more generic check for Office 2007+ files and files protected via third-party tools different to Microsoft Office.
 */
async function isPasswordProtected(file: File): Promise<boolean> {
    const cfb = parseCFB(new Uint8Array(await file.arrayBuffer()), { type: 'buffer' });

    const target = cfb.FileIndex.find(({name}) => {
        switch (name) {
            case 'EncryptionInfo':
            case 'EncryptedPackage':
                return true;
            default:
                return false;
        }
    });

    if (target) {
        return true;
    }

    return false;
}

/**
 * Seek encryption flags in a .docx/.xlsx/.pptx file (OOXML - zip based).
 */
async function isZipWithEncryptedInfo(file: File): Promise<boolean> {
    const text = await file.text();

    return text.includes('Encryption') || text.includes('EncryptedPackage');
}

/**
 * Return the first `length` bytes of the given file.
 */
async function readFileHeader(file: Blob, length: number): Promise<Uint8Array> {
    const buffer = await file.slice(0, length).arrayBuffer();

    return new Uint8Array(buffer);
}

/**
 * Return whether the file MIME is `application/msword` / `application/x-msword` or, in non-strict mode, the file name ends with `.doc`.
 */
export function isDoc(file: File, strict = false): boolean {
    switch (true) {
        case file.type === 'application/msword':
        case file.type === 'application/x-msword':
            return true;
        case !strict && file.name.toLowerCase().endsWith('.doc'):
            return true;
        default:
            return false;
    }
}

/**
 * Return whether the file MIME is `application/vnd.openxmlformats-officedocument.wordprocessingml.document` or, in non-strict mode, the file name ends with `.docx`.
 */
export function isDocx(file: File, strict = false): boolean {
    switch (true) {
        case file.type === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            return true;
        case !strict && file.name.toLowerCase().endsWith('.docx'):
            return true;
        default:
            return false;
    }
}

/**
 * Return whether the file MIME is `application/vnd.ms-excel` or, in non-strict mode, the file name ends with `.xls`.
 *
 * Note xlt and xla files might have use same MIME type.
 */
export function isXls(file: File, strict = false): boolean {
    switch (true) {
        case file.type === 'application/vnd.ms-excel':
            return true;
        case !strict && file.name.toLowerCase().endsWith('.xls'):
            return true;
        default:
            return false;
    }
}

/**
 * Return whether the file MIME is `application/vnd.openxmlformats-officedocument.spreadsheetml.sheet` or, in non-strict mode, the file name ends with `.xlsx`.
 */
export function isXlsx(file: File, strict = false): boolean {
    switch (true) {
        case file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            return true;
        case !strict && file.name.toLowerCase().endsWith('.xlsx'):
            return true;
        default:
            return false;
    }
}

/**
 * Return whether the file MIME is `application/vnd.ms-powerpoint` or, in non-strict mode, the file name ends with `.ppt`.
 *
 * Note pot, pps, and ppa files might use the same MIME type.
 */
export function isPpt(file: File, strict = false): boolean {
    switch (true) {
        case file.type === 'application/vnd.ms-powerpoint':
            return true;
        case !strict && file.name.toLowerCase().endsWith('.ppt'):
            return true;
        default:
            return false;
    }
}

/**
 * Return whether the file MIME is `application/vnd.openxmlformats-officedocument.presentationml.presentation` or, in non-strict mode, the file name ends with `.pptx`.
 */
export function isPptx(file: File, strict = false): boolean {
    switch (true) {
        case file.type === 'application/vnd.openxmlformats-officedocument.presentationml.presentation':
            return true;
        case !strict && file.name.toLowerCase().endsWith('.pptx'):
            return true;
        default:
            return false;
    }
}

/**
 * Return whether the file MIME is `application/pdf` / `application-x-pdf` or, in non-strict mode, the file name ends with `.pdf`.
 */
export function isPDF(file: File, strict = false): boolean {
    switch (true) {
        case file.type === 'application/pdf':
        case file.type === 'application/x-pdf':
            return true;
        case !strict && file.name.toLowerCase().endsWith('.pdf'):
            return true;
        default:
            return false;
    }
}

/**
 * Return whether the provided doc file is password protected (OLE2 format).
 */
export async function isDocPasswordProtected(file: File): Promise<boolean> {
    const header = await readFileHeader(file, 512);

    // Reading header[0x0B] requires a minimum file length of 12 bytes.
    // It's safe in practice, but for correctness and resilience, return early if the file is too short.
    if (header.length < 0x0B) {
        return false;
    }

    // Encrypted .doc files have specific flags in the FIB (File Information Block)
    // 0x0B bit (bit 0 of byte 0x0B) of the FIB base must be set.
    if (header[0x0B] & 0x01) {
        return true;
    }

    // Previous check only works if the FIB starts at a known offset, which in many .doc files is 512 bytes, but not always.
    // When the file doesn’t place the FIB at 0x00 — it can be embedded in a stream (WordDocument) inside the compound file,
    // or within a CDFV2 Encrypted envelope.
    return await isPasswordProtected(file);
}

/**
 * Return whether the provided docx file is password protected.
 *
 * When `strict` mode is disabled (default), legacy formats will be scanned as well (e.g., misnamed .doc files).
 */
export async function isDocxPasswordProtected(file: File, strict = false): Promise<boolean> {
    if (!strict && ! (await isOOXML(file))) {
        // Possibly misnamed file, treat it as legacy doc format.
        return isDocPasswordProtected(file);
    }

    return isZipWithEncryptedInfo(file);
}

/**
 * Return whether the provided xlsx file is password protected.
 *
 * When `strict` mode is disabled (default), legacy formats will be scanned as well (e.g., misnamed .xls files).
 */
export async function isXlsxPasswordProtected(file: File, strict = false): Promise<boolean> {
    if (!strict && ! (await isOOXML(file))) {
        // Possibly misnamed file, treat it as legacy doc format.
        return isXlsPasswordProtected(file);
    }

    return isZipWithEncryptedInfo(file);
}

/**
 * Return whether the provided pptx file is password protected.
 *
 * When `strict` mode is disabled (default), legacy formats will be scanned as well (e.g., misnamed .ppt files).
 */
export async function isPptxPasswordProtected(file: File, strict = false): Promise<boolean> {
    if (!strict && ! (await isOOXML(file))) {
        // Possibly misnamed file, treat it as legacy doc format.
        return isPptPasswordProtected(file);
    }

    return isZipWithEncryptedInfo(file);
}

/**
 * Return whether the provided xls file is password protected (BIFF format).
 */
export async function isXlsPasswordProtected(file: File): Promise<boolean> {
    const header = await readFileHeader(file, 1024);

    for (let i = 0; i < header.length - 1; i++) {
        // Protection indicated by the presence of the "FilePass" record (0x2F00).
        if (header[i] === 0x2F && header[i + 1] === 0x00) {
            return true;
        }
    }

    return false;
}

/**
 * Return whether the provided ppt file is password protected.
 */
export async function isPptPasswordProtected(file: File): Promise<boolean> {
    const header = await readFileHeader(file, 512);

    for (let i = 0; i < header.length - 1; i++) {
        // Similar heuristic to `isXlsPasswordProtected()`.
        if (header[i] === 0x2F && header[i + 1] === 0x00) {
            return true;
        }
    }

    return false;
}

/**
 * Return whether the provided pdf file is password protected.
 */
export async function isPDFPasswordProtected(file: File): Promise<boolean> {
    const header = await readFileHeader(file, 8192);
    const text = new TextDecoder().decode(header);

    // Look for "/Encrypt" followed by an indirect object reference (e.g., " 5 0 R") with:
    // const encryptMatch = text.match(/\/Encrypt\s+\d+\s+\d+\s+R/);
    // Tighten the regex to ensure /Encrypt is part of a dictionary key in the trailer, to avoid false positives.
    const encryptMatch = text.match(/<<[^>]*\/Encrypt\s+\d+\s+\d+\s+R[^>]*>>/);

    return !!encryptMatch;
}
