
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

pub struct Ods {
    path: PathBuf,
    data: Cursor<String>
}

impl HasKind for Ods {
    fn kind(&self) -> &'static str {
        "Open Office Spreadsheet"
    }
    
    fn ext(&self) -> &'static str {
        "ods"
    }
}

impl OpenOfficeDoc<Ods> for Ods {

    fn open<P: AsRef<Path>>(path: P) -> io::Result<Ods> {
        
        let text = doc::open_doc_read_data(path.as_ref(), "content.xml", &["text:p"])?;

        Ok(
            Ods {
                path: path.as_ref().to_path_buf(),
                data: Cursor::new(text)
            }
        )
    }
}

impl Read for Ods {
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
        let _ = Ods::open(Path::new("samples/sample.ods"));
    }

    #[test]
    fn read(){
        let mut f = Ods::open(Path::new("samples/sample.ods")).unwrap();

        let mut data = String::new();
        let len = f.read_to_string(&mut data).unwrap();
        println!("len: {}, data: {}", len, data);
    }
}
