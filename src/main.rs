use std::env::args;
use std::path::Path;
use docx_you_want as dyw;
use docx_you_want::Error;
use std::process::exit;


fn main() {
    let args: Vec<_> = args().collect();
    if args.len() != 3 {
        println!("1st arg: PDF file, 2nd arg: .docx file");
        exit(-1)
    }
    let src = Path::new(&args[1]);
    let dst = Path::new(&args[2]);
    if let Err(e) = convert(src, dst) {
        let msg = match e {
            Error::IoError => "An error occurred during I/O.",
            Error::ImageError => "Something went wrong while processing the images.",
            Error::InkscapeNotFound => "Inkscape not found. Consider installing inkscape?",
            Error::PDFInvalid => "Invalid PDF."
        };
        eprint!("{}", msg);
        exit(-1);
    }
}

fn convert(src: &Path, dst: &Path) -> dyw::Result<()> {
    let mut docx = dyw::Docx::new()?;
    docx.convert_pdf(src)?;
    docx.generate_docx(&dst.to_owned())?;
    Ok(())
}
