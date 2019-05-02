use zip::ZipArchive;

use xml::events::Event;
use xml::reader::Reader;

use std::clone::Clone;
use std::fs::File;
use std::io;
use std::io::prelude::*;
use std::io::Cursor;
use std::path::{Path, PathBuf};
use zip::read::ZipFile;

use doc;
use doc::{HasKind, OpenOfficeDoc};

pub struct Odt {
    path: PathBuf,
    data: Cursor<String>,
}

impl HasKind for Odt {
    fn kind(&self) -> &'static str {
        "Open Office Document"
    }

    fn ext(&self) -> &'static str {
        "odt"
    }
}

impl OpenOfficeDoc<Odt> for Odt {
    fn open<P: AsRef<Path>>(path: P) -> io::Result<Odt> {
        let text = doc::open_doc_read_data(path.as_ref(), "content.xml", &["text:p"])?;

        Ok(Odt {
            path: path.as_ref().to_path_buf(),
            data: Cursor::new(text),
        })
    }
}

impl Read for Odt {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        self.data.read(buf)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::path::{Path, PathBuf};

    #[test]
    fn instantiate() {
        let _ = Odt::open(Path::new("samples/sample.odt"));
    }

    #[test]
    fn read() {
        let mut f = Odt::open(Path::new("samples/sample.odt")).unwrap();

        let mut data = String::new();
        let len = f.read_to_string(&mut data).unwrap();
        println!("len: {}, data: {}", len, data);
    }
}
