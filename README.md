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
cd ~/working
git clone https://github.com/ptsefton/pptx_to_md.git
cd pptx_to_md
```
uv run pptx2md.py your-preso.pptx
```

### Usage

To create a markdown version of a `.pptx` file, run this:

    uv run pptx2md.py  your-preso.pptx 

The result ia directory `your-preso/` and with an `index.md` file in it.

The script WILL NOT create images for the slides (but see below the experimental part) - for that you need a PDF file. The best way to get a PDF is to manually create it from the application you used to make the .pptx file, usually Microsoft Word or Google Slides. if you have a PDF file with the same name as the .pprtx with a .pdf extension run:

    uv run pptx2md.py  your-preso.pptx 

If you're using Pelican or another CMS that needs it you can also add a prefix to the image paths use the `-i` flag:

    uv run pptx2md.py  -i {attach} your-preso.pptx 

### Experimental

If you would like to try your luck at getting Libre/Open Office to generate the PDF for you you can do this:

```
# On a mac with LibreOffice installed

uv run pptx2md.py --soffice /Applications/LibreOffice.app/Contents/MacOS/soffice  --topdf  preso.pptx

# On linux or where soffice is in your path
uv run pptx2md.py --topdf  preso.pptx
```

### As a docker container (experimental)

Build the container (this image installs `uv` and the Python dependencies):

```bash
docker build -t pptx2md .
```

Run the container (mount current directory as /data):

```bash
docker run --rm -v "${PWD}":/data pptx2md yourfile.pptx
```







