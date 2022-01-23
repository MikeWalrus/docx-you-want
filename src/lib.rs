/* This file is part of docx-you-want.

   docx-you-want is free software: you can redistribute it and/or modify
   it under the terms of the GNU General Public License as published by
   the Free Software Foundation, either version 3 of the License, or
   (at your option) any later version.

   docx-you-want is distributed in the hope that it will be useful,
   but WITHOUT ANY WARRANTY; without even the implied warranty of
   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
   GNU General Public License for more details.

   You should have received a copy of the GNU General Public License
   along with docx-you-want.  If not, see <https://www.gnu.org/licenses/>.
*/

#![recursion_limit = "512"]

use std::ffi::OsStr;
use std::fs::{copy, read_to_string, remove_file, write};
use std::io::{self, ErrorKind, Write};
use std::path::{Path, PathBuf};
use std::process::Command;
use tempfile::TempDir;
use zip_extensions::zip_create_from_directory;

#[derive(Debug)]
pub enum Error {
    IoError,
    ImageError,
    InkscapeNotFound,
    PDFInvalid,
}

pub type Result<T> = std::result::Result<T, Error>;

impl From<std::io::Error> for Error {
    fn from(_: std::io::Error) -> Error {
        Error::IoError
    }
}

impl From<usvg::Error> for Error {
    fn from(_: usvg::Error) -> Error {
        Error::ImageError
    }
}

impl From<png::EncodingError> for Error {
    fn from(_: png::EncodingError) -> Error {
        Error::ImageError
    }
}

impl From<zip::result::ZipError> for Error {
    fn from(_: zip::result::ZipError) -> Error {
        Error::IoError
    }
}

fn px_to_emu(px: f64) -> i32 {
    let dpi = 96.0;
    let emus_per_inch = 914400.0;
    (px / dpi * emus_per_inch) as i32
}

fn px_to_twenties_of_pt(px: f64) -> i32 {
    let dpi = 96.0;
    let pt_per_inch = 72.0;
    (px / dpi * pt_per_inch * 20.0) as i32
}

fn get_filename(svg: &Path) -> &str {
    svg.file_name().unwrap().to_str().unwrap()
}

fn read_svg(src: &Path) -> Result<usvg::Tree> {
    let opt = usvg::Options::default();
    let svg_data = std::fs::read(src)?;
    Ok(usvg::Tree::from_data(&svg_data, &opt)?)
}

fn save_png(dst: &Path, rtree: &usvg::Tree) -> Result<()> {
    let size = rtree.svg_node().size.to_screen_size();
    let mut pixmap = tiny_skia::Pixmap::new(size.width(), size.height()).unwrap();
    resvg::render(rtree, usvg::FitTo::Original, pixmap.as_mut()).ok_or(Error::ImageError)?;
    pixmap.save_png(dst)?;
    Ok(())
}

fn get_png_path(prefix: &Path, svg_path: &Path) -> Result<PathBuf> {
    let filename = svg_path
        .file_name()
        .unwrap()
        .to_str()
        .ok_or(Error::IoError)?
        .replace("svg", "png");
    Ok(prefix.join(Path::new(&filename)))
}

pub struct Docx {
    dir: TempDir,
    media_dir: PathBuf,
    doc: PathBuf,
    rels: PathBuf,
    next_id: i32,
    doc_string: String,
    rels_string: String,
    size: usvg::Size,
}

impl Docx {
    pub fn new() -> Result<Docx> {
        let dir = TempDir::new()?;
        Docx::copy_base_files(&dir)?;
        let path = dir.path();
        let doc: PathBuf = [path.as_os_str(), OsStr::new("word/document.xml")]
            .iter()
            .collect();
        let rels: PathBuf = [path.as_os_str(), OsStr::new("word/_rels/document.xml.rels")]
            .iter()
            .collect();
        let media_dir = [path.as_os_str(), OsStr::new("word/media")]
            .iter()
            .collect();
        Ok(Docx {
            dir,
            media_dir,
            doc,
            rels,
            next_id: 0,
            doc_string: String::new(),
            rels_string: String::new(),
            size: usvg::Size::new(793.707, 1122.52).unwrap(),
        })
    }

    fn copy_base_files(dir: &TempDir) -> Result<()> {
        let fixtures_zip = include_bytes!("../fixtures/fixtures.zip");
        let mut zip_path = dir.path().to_owned();
        zip_path.push("fixtures.zip");
        std::fs::write(&zip_path, fixtures_zip)?;
        zip_extensions::read::zip_extract(&zip_path, &dir.path().to_owned())?;
        remove_file(zip_path)?;
        Ok(())
    }

    fn add_image_svg(&mut self, svg: &Path) -> Result<()> {
        let tree = read_svg(svg)?;
        let png = get_png_path(&self.media_dir, svg)?;
        save_png(&png, &tree)?;
        let svg_copy = &self
            .media_dir
            .join(Path::new(svg.file_name().ok_or(Error::IoError)?));
        if svg != svg_copy {
            copy(svg, svg_copy)?;
        }
        self.add_to_doc(svg_copy, &png, &tree.svg_node().size);
        print!(".");
        io::stdout().flush()?;
        Ok(())
    }

    fn next_id(&mut self) -> i32 {
        let ret = self.next_id;
        self.next_id += 1;
        ret
    }

    fn add_to_doc(&mut self, svg: &Path, png: &Path, size: &usvg::Size) {
        let svg_id = self.next_id();
        let png_id = self.next_id();
        let svg_rid = format!("rId{}", svg_id);
        let png_rid = format!("rId{}", png_id);
        let width = px_to_emu(size.width());
        let height = px_to_emu(size.height());
        self.doc_string = format!(
            "{}{}",
            self.doc_string,
            format_xml::xml! {
              <w:p>
                <w:pPr>
                    <w:widowControl/>
                    <w:jc w:val="left"/>
                </w:pPr>
                <w:r>
                    <w:rPr>
                        <w:noProof/>
                    </w:rPr>
                    <w:drawing>
                        <wp:inline distT="0" distB="0" distL="0" distR="0">
                            <wp:extent cx={width} cy={height}/>
                            <wp:effectExtent l="0" t="0" r="0" b="0"/>
                            <wp:docPr id={svg_id} name={svg_id}/>
                            <wp:cNvGraphicFramePr>
                                <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                            </wp:cNvGraphicFramePr>
                            <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                    <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                        <pic:nvPicPr>
                                            <pic:cNvPr id="1" name=""/>
                                            <pic:cNvPicPr/>
                                        </pic:nvPicPr>
                                        <pic:blipFill>
                                            <a:blip r:embed={png_rid}>
                                                <a:extLst>
                                                    <a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">
                                                        <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed={svg_rid}/>
                                                    </a:ext>
                                                </a:extLst>
                                            </a:blip>
                                            <a:stretch>
                                                <a:fillRect/>
                                            </a:stretch>
                                        </pic:blipFill>
                                        <pic:spPr>
                                            <a:xfrm>
                                                <a:off x="0" y="0"/>
                                                <a:ext cx={width} cy={height}/>
                                            </a:xfrm>
                                            <a:prstGeom prst="rect">
                                                <a:avLst/>
                                            </a:prstGeom>
                                        </pic:spPr>
                                    </pic:pic>
                                </a:graphicData>
                            </a:graphic>
                        </wp:inline>
                    </w:drawing>
                </w:r>
              </w:p>
            }
        );
        self.add_relationship(&svg_rid, get_filename(svg));
        self.add_relationship(&png_rid, get_filename(png))
    }

    fn add_relationship(&mut self, rid: &str, filename: &str) {
        let target = format!("media/{}", filename);
        self.rels_string = format!(
            "{}{}",
            self.rels_string,
            format_xml::xml! {
                <Relationship Id={rid} Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target={target}/>
            }
        )
    }

    pub fn generate_docx(self, p: &PathBuf) -> Result<()> {
        self.write_to_files()?;
        zip_create_from_directory(p, &PathBuf::from(self.dir.path()))?;
        Ok(())
    }

    fn write_to_files(&self) -> Result<()> {
        Docx::insert_in_file(&self.doc, &self.doc_string)?;
        Docx::insert_in_file(&self.rels, &self.rels_string)?;
        self.change_size()?;
        Ok(())
    }

    fn change_size(&self) -> Result<()> {
        let s = read_to_string(&self.doc)?
            .replace(
                "!WIDTH!",
                &px_to_twenties_of_pt(self.size.width()).to_string(),
            )
            .replace(
                "!HEIGHT!",
                &px_to_twenties_of_pt(self.size.height()).to_string(),
            );
        write(&self.doc, s)?;
        Ok(())
    }

    fn insert_in_file(path: &Path, content: &str) -> Result<()> {
        let s = read_to_string(path)?.replace("!INSERT_HERE!", content);
        write(path, s)?;
        Ok(())
    }

    pub fn convert_pdf(&mut self, pdf: &Path) -> Result<()> {
        let mut page = 0;
        let mut images: Vec<PathBuf> = Vec::new();
        print!("Calling Inkscape to generate images ");
        loop {
            page += 1;
            let image = PathBuf::from(&self.media_dir).join(format! {"{}.svg", page});
            let output = match Command::new("inkscape")
                .arg(pdf)
                .arg(format!("--pdf-page={}", page))
                .arg("-o")
                .arg(&image)
                .arg("--pdf-poppler")
                .output()
            {
                Err(e) => {
                    return if let ErrorKind::NotFound = e.kind() {
                        Err(Error::InkscapeNotFound)
                    } else {
                        Err(Error::IoError)
                    };
                }
                Ok(output) => (output),
            };
            print!(".");
            io::stdout().flush()?;
            if output.stderr.is_empty() {
                images.push(image);
                continue;
            }
            remove_file(&image)?;
            println!(" Done.");
            break;
        }
        print!("Getting the size of the first page ... ");
        self.size = read_svg(images.get(0).ok_or(Error::PDFInvalid)?)?
            .svg_node()
            .size;
        println!("Done.");
        print!("Adding all the images ");
        io::stdout().flush()?;
        images.iter().try_for_each(|i| self.add_image_svg(i))
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::fs::read_dir;

    fn get_children(fixtures_dir: &Path) -> Result<Vec<PathBuf>> {
        let children: std::result::Result<Vec<_>, _> = read_dir(fixtures_dir)?.collect();
        let children: Vec<PathBuf> = children?.iter().map(|i| i.path()).collect();
        Ok(children)
    }

    #[test]
    fn test_dir() -> Result<()> {
        let docx = Docx::new().unwrap();
        let dir = docx.dir.path();
        assert!(dir.exists());
        let children = get_children(dir)?;
        let mut children_str: Vec<&str> = children
            .iter()
            .map(|i| i.file_name().unwrap().to_str().unwrap())
            .collect();
        let mut result = vec!["word", "[Content_Types].xml", "_rels"];

        children_str.sort_unstable();
        result.sort_unstable();
        assert_eq!(children_str, result);
        Ok(())
    }

    #[test]
    fn test_tmp_dir_drop() {
        let docx = Docx::new().unwrap();
        let dir = docx.dir.path();
        let dir_string = String::from(dir.to_str().unwrap());
        drop(docx);
        let should_be_deleted = Path::new(&dir_string);
        assert!(!should_be_deleted.exists());
    }

    fn get_test_svg() -> PathBuf {
        let tests_dir = String::from(env!("CARGO_MANIFEST_DIR")) + "/tests/";
        PathBuf::from(format!("{}{}", tests_dir, "2.svg"))
    }

    fn get_tests_dir() -> String {
        String::from(env!("CARGO_MANIFEST_DIR")) + "/tests/"
    }

    #[test]
    fn test_add_svg() {
        let mut docx = Docx::new().unwrap();
        docx.add_image_svg(&get_test_svg()).unwrap();
        assert_eq!(docx.doc_string,
                   format_xml::xml! {
<w:p>
    <w:pPr>
        <w:widowControl />
        <w:jc w:val="left" />
    </w:pPr>
    <w:r>
        <w:rPr>
            <w:noProof />
        </w:rPr>
        <w:drawing>
            <wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="7560000" cy="10692000" />
                <wp:effectExtent l="0" t="0" r="0" b="0" />
                <wp:docPr id="0" name="0" />
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1" />
                </wp:cNvGraphicFramePr>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:nvPicPr>
                                <pic:cNvPr id="1" name="" />
                                <pic:cNvPicPr />
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:embed="rId1">
                                    <a:extLst>
                                        <a:ext uri="{{96DAC541-7B7A-43D3-8B79-37D633B846F1}}">
                                            <asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId0" />
                                        </a:ext>
                                    </a:extLst>
                                </a:blip>
                                <a:stretch>
                                    <a:fillRect />
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0" />
                                    <a:ext cx="7560000" cy="10692000" />
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst />
                                </a:prstGeom>
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:inline>
        </w:drawing>
    </w:r>
</w:p>
            }.to_string());
        assert_eq!(docx.rels_string,
                   format_xml::xml! {
<Relationship Id="rId0" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/2.svg" />
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/2.png" />
            }.to_string())
    }

    #[test]
    fn test_write_to_file() {
        let mut docx = Docx::new().unwrap();
        docx.doc_string = String::from("<p></p>");
        docx.write_to_files().unwrap();
    }

    #[test]
    fn test_generate_docx() {
        let mut docx = Docx::new().unwrap();
        docx.add_image_svg(&get_test_svg()).unwrap();
        docx.generate_docx(&PathBuf::from(get_tests_dir() + "a.docx"))
            .unwrap();
    }

    #[test]
    fn test_size() {
        assert_eq!(px_to_twenties_of_pt(793.707), 11905)
    }
}
