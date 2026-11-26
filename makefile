# Build the static site from course material kept outside this repo.
# Override the source location on the command line if it moves, e.g.
#   make html SRC=/path/to/CitiesCourse
# Needs python-pptx (pip install python-pptx), pandoc, and Keynote for .ppt decks.
# Converted decks are cached in .deck-cache/; `make clean-cache` clears it.
SRC ?= $(HOME)/Documents/CoursesTaught/CitiesCourse

FORCE:

html: FORCE
	python3 bin/build.py --src "$(SRC)"

clean: FORCE
	rm -rf lectures handouts spreadsheets quizzes gallery index.html

clean-cache: FORCE
	rm -rf .deck-cache
