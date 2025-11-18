# pptx_to_md

Powerpoint to Markdown converter - Python 3 script

NOTE: This works on `.pptx` files ONLY - if you're using Google Drive or Keynote etc export your presentation to .pptx first.

What does it do:

Takes a PowerPoint file and a PDF of the same file (generated from the same application) and generates a Markdown version in a directory with an image for each slide and a yaml metadata block as used by various static site generators.

## Audience

This is for people who know how to run Python or Docker. Instructions are for MacOs. The examples below show the `uv` runner (replace with `python`/`python3` if you prefer).

## About 

This is a simple rough and ready way to turn a Powerpoint slide deck into Markdown.


## Installation and running with `uv`

This project no longer requires Poetry. Use the `uv` runner to execute the script directly.

1) Get the code:

```bash
git clone https://github.com/ptsefton/pptx_to_md.git
cd pptx_to_md
```

```
uv run pptx2md.py your-preso.pptx
```

### Usage

#### Quick Test

Try the tool with the included test file:

```bash
# Basic usage - creates a directory called 'test-preso' with index.md and SVG images
uv run pptx2md.py test_files/test-preso.pptx
```

The result is a directory `test-files/test-preso/ with an `index.md` file in it.


```bash
# Specify custom output directory and generate and index.html file
uv run pptx2md.py -d out --html test_files/test-preso.pptx
```


#### Available Options

**Custom output directory:**
```bash
# Specify where to write the output
uv run pptx2md.py -d /path/to/output your-preso.pptx
```

**Generate HTML:**
```bash
# Create both index.md and index.html
uv run pptx2md.py --html your-preso.pptx
```

**Image path prefix:**
```bash
# Add a prefix to image paths (useful for Pelican or other CMS)
uv run pptx2md.py -i {attach} your-preso.pptx
```

**Combine options:**
```bash
uv run pptx2md.py -d output --html -i /images/ your-preso.pptx
``` 
The script WILL NOT create images for the slides (but see below the experimental part) - for that you need a PDF file. The best way to get a PDF is to manually create it from the application you used to make the .pptx file, usually Microsoft Word or Google Slides. if you have a PDF file with the same name as the .pptx with a .pdf extension run:

    uv run pptx2md.py  your-preso.pptx 

If you're using Pelican or another CMS that needs it you can also add a prefix to the image paths use the `-i` flag:

    uv run pptx2md.py  -i {attach} your-preso.pptx 

### Experimental  (Have not checked this as of 2025-11-19)

If you would like to try your luck at getting Libre/Open Office to generate the PDF for you you can do this:

```bash
# On a mac with LibreOffice installed

uv run pptx2md.py --soffice /Applications/LibreOffice.app/Contents/MacOS/soffice  --topdf  preso.pptx

# On linux or where soffice is in your path
uv run pptx2md.py --topdf  preso.pptx
```

### As a docker container 

Build the container (this image installs `uv` and the Python dependencies):

```bash
docker build -t pptx2md .
```

Run the container (mount current directory as /data):

```bash
docker run --rm -v "${PWD}":/data pptx2md yourfile.pptx
```







