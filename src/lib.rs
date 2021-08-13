use tempfile::TempDir;
use fs_extra::dir::CopyOptions;
use std::path::{Path, PathBuf};
use std::fs::read_dir;

#[derive(Debug)]
enum Error {
    IoError
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

struct Docx {
    dir: TempDir,
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
        Ok(Docx { dir })
    }

    fn copy_base_files(dir: &TempDir) -> Result<()> {
        let fixtures_dir = Path::new("/home/mike/repos/rust/docx-you-want/fixtures");
        let children = get_children(fixtures_dir)?;
        fs_extra::copy_items(&children, &dir, &CopyOptions::new())?;
        Ok(())
    }
}


#[cfg(test)]
mod test {
    use super::*;

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
        let children = get_children(dir)?;
        let children_str: Vec<&str> = children
            .iter()
            .map(|i| i.file_name().unwrap().to_str().unwrap())
            .collect();
        assert_eq!(children_str, vec!["word", "[Content_Types].xml", "_rels"]);
        Ok(())
    }
}