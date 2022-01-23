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

use docx_you_want as dyw;
use docx_you_want::Error;
use std::env::args;
use std::io::{self, Write};
use std::path::Path;
use std::process::exit;

fn main() {
    let args: Vec<_> = args().collect();
    if args.len() != 3 {
        println!(
            "Usage: {} <path to PDF> <path to result DOCX file>",
            args[0]
        );
        exit(-1)
    }
    let src = Path::new(&args[1]);
    let dst = Path::new(&args[2]);
    if let Err(e) = convert(src, dst) {
        let msg = match e {
            Error::IoError => "An error occurred during I/O.",
            Error::ImageError => "Something went wrong while processing the images.",
            Error::InkscapeNotFound => "Inkscape not found. Consider installing inkscape?",
            Error::PDFInvalid => "Invalid PDF.",
        };
        eprint!("{}", msg);
        exit(-1);
    }
}

fn convert(src: &Path, dst: &Path) -> dyw::Result<()> {
    let mut docx = dyw::Docx::new()?;
    docx.convert_pdf(src)?;
    println!("Done");
    print!("Generating the final result ... ");
    io::stdout().flush()?;
    docx.generate_docx(&dst.to_owned())?;
    println!("Done.");
    Ok(())
}
