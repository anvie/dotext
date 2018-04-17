
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
        let file = File::open(path.as_ref())?;
        let mut archive = ZipArchive::new(file)?;

        let mut xml_data = String::new();

        for i in 0..archive.len(){
            let mut c_file = archive.by_index(i).unwrap();
            if c_file.name() == "content.xml" {
                c_file.read_to_string(&mut xml_data);
                break
            }
        }

        let mut xml_reader = Reader::from_str(xml_data.as_ref());

        let mut buf = Vec::new();
        let mut txt = Vec::new();

        if xml_data.len() > 0 {
            let mut to_read = false;
            loop {
                match xml_reader.read_event(&mut buf){
                    Ok(Event::Start(ref e)) => {
                        match e.name() {
                            b"text:p" => {
                                to_read = true;
                                txt.push("\n\n".to_string());
                            },
                            _ => (),
                        }
                    },
                    Ok(Event::Text(e)) => {
                        if to_read {
                            txt.push(e.unescape_and_decode(&xml_reader).unwrap());
                            to_read = false;
                        }
                    },
                    Ok(Event::Eof) => break,
                    Err(e) => panic!("Error at position {}: {:?}", xml_reader.buffer_position(), e),
                    _ => (),
                }
            }
        }

        Ok(
            Odt {
                path: path.as_ref().to_path_buf(),
                data: Cursor::new(txt.join(""))
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
