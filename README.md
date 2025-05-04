# pwcheckr 0.0.1

Weak detection of certain file types and usage of password/encryption in them.

Supported files:

- `.doc` and `.docx`
- `.ppt` and `.pptx`
- `.xls` and `.xlsx`
- `.pdf`

This is a beta version, it might return false negatives, but **never** false positives (except for carefully crafted inputted files).

## File type detection

Detection is first done on a file's mime, but always falls back to checking the extension (case-insensitive).

```ts
// E.g.: a file with a mime type of `application/vnd.openxmlformats-officedocument.wordprocessingml.document`, with `doc` extension.
const file = document.querySelector('input[type=file]').files[0]; // ie: "/path/to/docx/renamed_to.doc";

// It will be detected as a `.doc` as well as a `.docx` file.
isDoc(file)  // true
isDocx(file) // true
```

## Test fixtures

The password used to encrypt all protected files is `password`.

### Empty files

These files have zero byte size and are used to test the detection of empty files.

- `empty.doc`
- `empty.docx`
- `empty.pdf`
- `empty.ppt`
- `empty.pptx`
- `empty.xls`
- `empty.xlsx`

### Documents

- `google-protected.doc` - A docx file, protected, authored via docs.google.com.
- `google-unprotected.docx` - A docx file, unprotected, authored via docs.google.com.

### Presentations

- [ ] To do

### Spreadsheets

- [ ] To do