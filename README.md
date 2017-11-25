Document File Text Extractor
=============================

[![Build Status](https://travis-ci.org/anvie/fred.svg?branch=master)](https://travis-ci.org/anvie/fred)

Simple Rust library to extract readable text from specific document format like Word Document (docx).
Currently only support several format, other format coming soon.

Supported Document
-------------------------


- [x] Microsoft Word (docx)
- [x] Microsoft Excel (xlsx)
- [x] Microsoft Power Point (pptx)
- [ ] OpenOffice Writer (odt)



Usage
------

```rust
let mut file = Docx::open("data/sample.docx").unwrap();
let mut isi = String::new();
let _ = file.read_to_string(&mut isi);
println!("CONTENT:");
println!("----------BEGIN----------");
println!("{}", isi);
println!("----------EOF----------");
```

Test
-----

```bash
$ cargo test
```

or run example:

```bash
$ cargo run --example readdocx data/sample.docx
```

[] Robin Sy.

