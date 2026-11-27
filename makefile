# Build the static site. Handouts, quizzes and photos come from course material
# kept outside this repo; lecture slides come from the saved copies in slides/.
# Override the source location on the command line if it moves, e.g.
#   make html SRC=/path/to/CitiesCourse
# Needs pandoc and Pillow; `make extract` also needs python-pptx and, for
# .ppt/.key decks, Keynote (converted decks are cached in .deck-cache/).
SRC ?= $(HOME)/Documents/CoursesTaught/CitiesCourse

FORCE:

html: FORCE
	python3 bin/build.py --src "$(SRC)"

# Re-read the lecture decks into slides/ and lectures/<deck>/ images.
# Overwrites any hand edits to slides/*.html.
extract: FORCE
	python3 bin/build.py --src "$(SRC)" --extract

# Redraw graphics/og-image.jpg, the preview shown when a link to the site is shared.
og-image: FORCE
	python3 bin/og_image.py

# Removes generated pages only; saved slides and their images are kept.
clean: FORCE
	rm -rf lectures/*.html handouts spreadsheets quizzes gallery index.html

clean-cache: FORCE
	rm -rf .deck-cache
