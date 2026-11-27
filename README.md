# Cities

Material for my course on cities: *Cities, Culture and Economy* ("What Is a City?").

Course site: https://gcallah.github.io/Cities/

## Building the site

Handouts (`.txt`), quizzes (`.docx`) and photographs of Siena come from course
material kept outside this repo, in `~/Documents/CoursesTaught/CitiesCourse`.
Lecture slides were read from the decks there once and saved in the repo:
the text of each deck in `slides/<deck>.html`, its images in `lectures/<deck>/`.

- `make html` rebuilds the site (`index.html`, `lectures/*.html`, `handouts/`,
  `quizzes/`, `gallery/`) from those saved slides; it never opens the decks.
  To fix a slide, edit `slides/<deck>.html` (each slide's title is its `<h2>`)
  and run `make html`.
- `make extract` re-reads the decks and **overwrites** `slides/` and the slide
  images, losing any hand edits. Needs `python-pptx` and, for old `.ppt`/`.key`
  decks, Keynote (conversions are cached in `.deck-cache/`, gitignored).
- `make og-image` redraws `graphics/og-image.jpg`, the preview picture shown when
  a link to the site is shared, from the backdrop painting, the site title and the
  theme colours. Run it after changing any of those. Every page links to it in its
  Open Graph tags, which use the absolute address `SITE_URL` in `bin/build.py`.

Override the source location with `make html SRC=...`. Requires `pandoc` and
Pillow (`pip install pillow python-pptx`).

What goes on the site, and in what order, is set by the tables at the top of
`bin/build.py`: `DECKS`, `HANDOUTS`, `SPREADSHEETS`, `QUIZZES`, `GALLERY` and the
`TOPICS` outline of the course. Colours, fonts and widths are set once, in
`:root` at the top of `style.css`.
