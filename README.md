Document File Reader
=======================

Simple Rust library to extract readable text from specific document format like Word Document (docx).
Currently only support for docx, other format coming soon.


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
$ cargo run --example readdocx data/filosofi-logo.docx
```

[] Robin Sy.

