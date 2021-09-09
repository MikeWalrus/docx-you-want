# docx-you-want
`docx-you-want` is a tool to convert a PDF document into a `.docx` file ... in an unusual way.
Since these two formats are inherently different, it is impossible to get a `.docx` file from a PDF without a noticeable difference in their appearances.
`docx-you-want` on the other hand, sort of preserve the look of the original PDF.

## What does it really do?
1. It calls [Inkscape](https://inkscape.org/) to convert every individual page of the PDF into SVGs, thus preserving its look.
   This means to run it, `inkscape` should be installed and in your `PATH`.
2. Then it inserts those images into a minimal `.docx` file, adding a PNG version of each also so that programs that don't support SVG in a `.docx` file have something to fall back on.
3. Finally, it zips the files and gives you the `.docx` (you want?).

## When to use this tool?
Hopefully never.

However, if someone asks you to send them a `.docx` version of your document and refuses to accept the PDF version that you only have, consider using it.
The next thing should be the person being very sad about your fake `.docx` document and wondering: is this *really* the `.docx` he wants?

## Why Rust?
My bad.

I really should have written it in bash or Python, none of which, including Rust, I am good at, though.
