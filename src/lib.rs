#![recursion_limit = "512"]

use tempfile::{TempDir, tempdir};
use fs_extra::dir::CopyOptions;
use std::path::{Path, PathBuf};
use std::fs::{read_dir, File, copy};
use std::ffi::OsStr;
use usvg;
use std::io::Read;
use usvg::Size;

#[derive(Debug)]
enum Error {
    IoError,
    ImageError,
}

type Result<T> = std::result::Result<T, Error>;

impl From<std::io::Error> for Error {
    fn from(_: std::io::Error) -> Error {
        Error::IoError
    }
}

impl From<fs_extra::error::Error> for Error {
    fn from(_: fs_extra::error::Error) -> Error {
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

fn px_to_emu(px: f64) -> i32 {
    let dpi = 96.0;
    let emus_per_inch = 914400.0;
    (px / dpi * emus_per_inch) as i32
}

fn get_filename(svg: &Path) -> &str {
    svg.file_name().unwrap().to_str().unwrap()
}

fn svg_to_png(src: &Path, dst: &Path) -> Result<()> {
    let rtree = read_svg(src)?;
    let size = rtree.svg_node().size;
    save_png(dst, &rtree);
    Ok(())
}

fn read_svg(src: &Path) -> Result<usvg::Tree> {
    let opt = usvg::Options::default();
    let svg_data = std::fs::read(src)?;
    Ok(usvg::Tree::from_data(&svg_data, &opt)?)
}

fn save_png(dst: &Path, rtree: &usvg::Tree) -> Result<()> {
    let size = rtree.svg_node().size.to_screen_size();
    let mut pixmap = tiny_skia::Pixmap::new(size.width(), size.height()).unwrap();
    resvg::render(&rtree, usvg::FitTo::Original, pixmap.as_mut()).ok_or(Error::ImageError)?;
    pixmap.save_png(dst)?;
    Ok(())
}

fn get_png_path(prefix: &Path, svg_path: &Path) -> Result<PathBuf> {
    let filename = svg_path.file_name().unwrap().to_str().ok_or(Error::IoError)?
        .replace("svg", "png");
    Ok(prefix.join(Path::new(&filename)))
}

struct Docx {
    dir: TempDir,
    media_dir: PathBuf,
    doc: File,
    rels: File,
    next_id: i32,
    doc_string: String,
    rels_string: String,
}

fn get_children(fixtures_dir: &Path) -> Result<Vec<PathBuf>> {
    let children: std::result::Result<Vec<_>, _> = read_dir(fixtures_dir)?.collect();
    let children: Vec<PathBuf> = children?.iter().map(|i| i.path()).collect();
    Ok(children)
}

impl Docx {
    fn new() -> Result<Docx> {
        let dir = TempDir::new()?;
        Docx::copy_base_files(&dir)?;
        let path = dir.path();
        let doc_path: PathBuf = [path.as_os_str(), OsStr::new("word/document.xml")].iter().collect();
        let doc = File::open(doc_path)?;
        let rels_path: PathBuf = [path.as_os_str(), OsStr::new("word/_rels/document.xml.rels")].iter().collect();
        let relations = File::open(rels_path)?;
        let media_dir = [path.as_os_str(), OsStr::new("word/media")].iter().collect();
        Ok(Docx { dir, media_dir, doc, rels: relations, next_id: 0, doc_string: String::new(), rels_string: String::new() })
    }

    fn copy_base_files(dir: &TempDir) -> Result<()> {
        let fixtures_dir = Path::new("/home/mike/repos/rust/docx-you-want/fixtures");
        let children = get_children(fixtures_dir)?;
        fs_extra::copy_items(&children, &dir, &CopyOptions::new())?;
        Ok(())
    }

    fn add_images(&self, _images: Vec<PathBuf>) -> Result<()> {
        let png_dir = tempdir()?;
        let _png_path_prefix = png_dir.path();
        Ok(())
    }

    fn add_image_svg(&mut self, svg: &Path) -> Result<()> {
        let tree = read_svg(svg)?;
        let png = get_png_path(&self.media_dir, svg)?;
        save_png(&png, &tree);
        let svg_copy = &self.media_dir.join(Path::new(svg.file_name().ok_or(Error::IoError)?));
        copy(svg, svg_copy)?;
        self.add_to_doc(svg_copy, &png, &tree.svg_node().size);
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
        self.doc_string = format!("{}{}", self.doc_string, format_xml::xml! {
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
        });
        self.add_relationship(&svg_rid, get_filename(svg));
        self.add_relationship(&png_rid, get_filename(png))
    }

    fn add_relationship(&mut self, rid: &str, filename: &str) {
        let target = format!("media/{}", filename);
        self.rels_string = format!("{}{}", self.rels_string, format_xml::xml! {
            <Relationship Id={rid} Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target={target}/>
        })
    }
}


#[cfg(test)]
mod tests {
    use super::*;
    use std::fs::remove_file;

    #[test]
    fn test() {
        let document = std::fs::read_to_string("/home/mike/repos/rust/docx-you-want/fixtures/word/document.xml").unwrap();
        let root: minidom::Element = document.parse().unwrap();
        println!("{:#?}", root);
        println!("{}", String::from(&root));
    }

    #[test]
    fn test_dir() -> Result<()>
    {
        let docx = Docx::new().unwrap();
        let dir = docx.dir.path();
        assert!(dir.exists());
        let children = get_children(&dir)?;
        let children_str: Vec<&str> = children
            .iter()
            .map(|i| i.file_name().unwrap().to_str().unwrap())
            .collect();
        assert_eq!(children_str, vec!["word", "[Content_Types].xml", "_rels"]);
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

    #[test]
    fn test_svg_to_png() {
        let tests_dir = String::from(env!("CARGO_MANIFEST_DIR")) + "/tests/";
        let dst = PathBuf::from(format!("{}{}", tests_dir, "2.png"));
        remove_file(&dst).unwrap();
        svg_to_png(&PathBuf::from(format!("{}{}", tests_dir, "2.svg")), &dst).unwrap();
        assert!(dst.exists())
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
                <wp:extent cx="7560055" cy="10692003" />
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
                                    <a:ext cx="7560055" cy="10692003" />
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
}