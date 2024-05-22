# pptx_to_md

Powerpoint to Markdown converter - Python 3 script

NOTE: This works on `.pptx` files ONLY - if you're using Google Drive or Keynote etc export your presentation to .pptx first.

What does it do:

Takes a PowerPoint file and a PDF of the same file (generated from the same application) and generates a Markdown version in a directory with an image for each slide and a yaml metadata block as used by various static site generators.

## Audience

This is for people who know how to run Python or Docker. Instructions are for MacOs.

## About 

This is a simple rough and ready way to turn a Powerpoint slide deck into Markdown.


## Installation

This project uses the [Poetry](https://python-poetry.org/) package manager. 

* Install Poetry if you don't have it.

* Get this code :
   ```
   cd ~/working
   git clone https://github.com/ptsefton/pptx_to_md.git
   cd pptx_to_md
   ```

* Set up a virtual environment with Poetry
  ```
  poetry install # Installs the dependencies in a virtual environment
  poetry shell # Starts an shell so you can use python
  ```


### Usage

To create a markdown version of a `.pptx` file, run this:

    python pptx2md.py  your-preso.pptx 

The result ia directory `your-preso/` and with an `index.md` file in it.

The script WILL NOT create images for the slides (but see below the experimental part) - for that you need a PDF file. The best way to get a PDF is to manually create it from the application you used to make the .pptx file, usually Microsoft Word or Google Slides. if you have a PDF file with the same name as the .pprtx with a .pdf extension run:

    python pptx2md.py  --pdf your-preso.pptx 

If you're using Pelican or another CMS that needs it you can also add a prefix to the image paths use the `-i` flag:

    python pptx2md.py -p -i {attach} your-preso.pptx 

### Experimental

If you would like to try your luck at getting Libre/Open Office to generate the PDF for you you can do this:

```
# On a mac with LibreOffice installed
python3 pptx2md.py --soffice /Applications/LibreOffice.app/Contents/MacOS/soffice  --topdf  preso.pptx

# On linux or where soffice is in your path
python3 pptx2md.py --topdf  preso.pptx
```

### As a docker container (experimental)

Make a container: 

`docker build -t rocxl .`

Run the container:

`docker run -v ${pwd} pptx2md yourfile.pptx`







