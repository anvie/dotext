use zip::ZipArchive;

use xml::events::Event;
use xml::reader::Reader;

use std::clone::Clone;
use std::fs::File;
use std::io;
use std::io::prelude::*;
use std::path::{Path, PathBuf};
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

pub(crate) fn open_doc_read_data<P: AsRef<Path>>(
    path: P,
    content_name: &str,
    tags: &[&str],
) -> io::Result<String> {
    let file = File::open(path.as_ref())?;
    let mut archive = ZipArchive::new(file)?;

    let mut xml_data = String::new();

    for i in 0..archive.len() {
        let mut c_file = archive.by_index(i).unwrap();
        if c_file.name() == content_name {
            c_file.read_to_string(&mut xml_data);
            break;
        }
    }

    let mut xml_reader = Reader::from_str(xml_data.as_ref());

    let mut buf = Vec::new();
    let mut txt = Vec::new();

    if xml_data.len() > 0 {
        let mut to_read = false;
        loop {
            match xml_reader.read_event(&mut buf) {
                Ok(Event::Start(ref e)) => {
                    for tag in tags {
                        if e.name() == tag.as_bytes() {
                            to_read = true;
                            if e.name() == b"text:p" {
                                txt.push("\n\n".to_string());
                            }
                            break;
                        }
                    }
                }
                Ok(Event::Text(e)) => {
                    if to_read {
                        txt.push(e.unescape_and_decode(&xml_reader).unwrap());
                        to_read = false;
                    }
                }
                Ok(Event::Eof) => break,
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

    Ok(txt.join(""))
}
