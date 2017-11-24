

extern crate fred;

use fred::Docx;
use std::io::Read;

fn main(){
    let mut file = Docx::open("data/sample.docx").unwrap();
    let mut isi = String::new();
    let _ = file.read_to_string(&mut isi);
    println!("ISI:");
    println!("----------BEGIN----------");
    println!("{}", isi);
    println!("----------EOF----------");
}
