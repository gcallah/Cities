# Cities

Material for my course on cities: *What Is a City?* (Cities, Culture, and the Economy).

Course site: https://gcallah.github.io/Cities/

## Building the site

The site is generated from course material kept outside this repo, in
`~/Documents/CoursesTaught/CitiesCourse`: lecture decks (`.ppt`, `.pptx`), a
handout (`.txt`), quizzes (`.docx`) and photographs of Siena.

Run `make html` to regenerate `index.html`, `lectures/`, `handouts/`, `quizzes/`
and `gallery/`. Override the source location with `make html SRC=...`.
Requires `pandoc`, `python-pptx` and Pillow (`pip install python-pptx pillow`),
and, on macOS, Keynote: old `.ppt`/`.key` decks are converted to `.pptx` through
Keynote and cached in `.deck-cache/` (gitignored; `make clean-cache` clears it).
Slide images are scaled down and large photos re-encoded as JPEG.

What goes on the site, and in what order, is set by the tables at the top of
`bin/build.py`: `DECKS`, `HANDOUTS`, `SPREADSHEETS`, `QUIZZES`, `GALLERY` and the
`TOPICS` outline of the course. Colours, fonts and widths are set once, in
`:root` at the top of `style.css`.
