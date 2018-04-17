


use zip::ZipArchive;

use xml::reader::Reader;
use xml::events::Event;

use std::path::{Path, PathBuf};
use std::fs::File;
use std::io::prelude::*;
use std::io;
use std::clone::Clone;
use zip::read::ZipFile;

pub trait HasKind {

    // kind
    fn kind(&self) -> &'static str;

    // extension
    fn ext(&self) -> &'static str;
}

pub trait MsDoc<T>: Read + HasKind {
    fn open<P: AsRef<Path>>(path: P) -> io::Result<T>;
}

pub trait OpenOfficeDoc<T>: Read + HasKind {
    fn open<P: AsRef<Path>>(path: P) -> io::Result<T>;
}
