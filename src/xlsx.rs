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

use doc::{HasKind, MsDoc};

pub struct Xlsx {
    path: PathBuf,
    data: Cursor<String>,
}

impl HasKind for Xlsx {
    fn kind(&self) -> &'static str {
        "Excel"
    }

    fn ext(&self) -> &'static str {
        "xlsx"
    }
}

impl MsDoc<Xlsx> for Xlsx {
    fn open<P: AsRef<Path>>(path: P) -> io::Result<Xlsx> {
        let file = File::open(path.as_ref())?;
        let mut archive = ZipArchive::new(file)?;

        let mut xml_data = String::new();
        //        let xml_data_list = Vec::new();

        for i in 0..archive.len() {
            let mut c_file = archive.by_index(i).unwrap();
            if c_file.name() == "xl/sharedStrings.xml"
                || c_file.name().starts_with("xl/charts/")
                || (c_file.name().starts_with("xl/worksheets") && c_file.name().ends_with(".xml"))
            {
                let mut _buff = String::new();
                c_file.read_to_string(&mut _buff);
                xml_data += _buff.as_str();
            }
        }

        let mut buf = Vec::new();
        let mut txt = Vec::new();

        if xml_data.len() > 0 {
            let mut to_read = false;
            let mut xml_reader = Reader::from_str(xml_data.as_ref());
            loop {
                match xml_reader.read_event(&mut buf) {
                    Ok(Event::Start(ref e)) => match e.name() {
                        b"t" => {
                            to_read = true;
                            txt.push("\n".to_string());
                        }
                        b"a:t" => {
                            to_read = true;
                            txt.push("\n".to_string());
                        }
                        _ => (),
                    },
                    Ok(Event::Text(e)) => {
                        if to_read {
                            let text = e.unescape_and_decode(&xml_reader).unwrap();
                            txt.push(text);
                            to_read = false;
                        }
                    }
                    Ok(Event::Eof) => break, // exits the loop when reaching end of file
                    Err(e) => {
                        return Err(io::Error::new(
                            io::ErrorKind::Other,
                            format!(
                                "Error at position {}: {:?}",
                                xml_reader.buffer_position(),
                                e
                            ),
                        ))
                    }
                    _ => (),
                }
            }
        }

        Ok(Xlsx {
            path: path.as_ref().to_path_buf(),
            data: Cursor::new(txt.join("")),
        })
    }
}

impl Read for Xlsx {
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
        let _ = Xlsx::open(Path::new("samples/sample.xlsx"));
    }

    #[test]
    fn read() {
        let mut f = Xlsx::open(Path::new("samples/sample.xlsx")).unwrap();

        let mut data = String::new();
        let len = f.read_to_string(&mut data).unwrap();
        println!("len: {}, data: {}", len, data);
    }
}
