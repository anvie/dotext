
use zip::ZipArchive;

use xml::reader::Reader;
use xml::events::Event;

use std::path::{Path, PathBuf};
use std::fs::File;
use std::io::prelude::*;
use std::io::Cursor;
use std::io;
use std::clone::Clone;
use zip::read::ZipFile;

use doc;
use doc::{OpenOfficeDoc, HasKind};

pub struct Odt {
    path: PathBuf,
    data: Cursor<String>
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

        Ok(
            Odt {
                path: path.as_ref().to_path_buf(),
                data: Cursor::new(text)
            }
        )
    }
}

impl Read for Odt {
    fn read(&mut self, buf: &mut [u8]) -> io::Result<usize> {
        self.data.read(buf)
    }
}


#[cfg(test)]
mod tests {
    use std::path::{Path, PathBuf};
    use super::*;

    #[test]
    fn instantiate(){
        let _ = Odt::open(Path::new("samples/sample.odt"));
    }

    #[test]
    fn read(){
        let mut f = Odt::open(Path::new("samples/sample.odt")).unwrap();

        let mut data = String::new();
        let len = f.read_to_string(&mut data).unwrap();
        println!("len: {}, data: {}", len, data);
    }
}
