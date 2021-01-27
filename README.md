# pptx_to_md

Powerpoint to Markdown converter - Python 3 script


## Audience

This is for people who know how to run Python, instructions are for MacOs.

## About 

This is a simple rough and ready way to turn a Powerpoint slide deck into Markdown, assuming that the slides have first been exported as a series of PNG images.

Hint: Format your slide notes as Markdown/Commonmark for instant gratification

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

## Usage

NOTE: This works on `.pptx` files ONLY - if you're using Google Drive or Keynote etc export your presentation to .pptx first.

The first step is to save the slides as PNG images. You can try to do this automatically (see below) but it's actually just as fast to do it manually.

### Manually generate images

* Open the presentation.

* From the `File` menu select `Export...` and choose `PNG` in `File Format` dropdown - also select `Save Every Slide`.

*  Click `Replace`

* Click OK on the confirmation dialogue "Each slide in your presentation has been saved as a separate file in the folder /your/path/presentation".

You should now have a directory named for your pptx file minus the extension.

### Or,automatically generate images

If you're on a Mac, and you have Powerpoint, try running with the -p flag:

   python pptx2md.py -p your-preso.pptx 

The first time you access a file, the operating system will make you do an authentication step.

*  If the script fails try running it again. 

*  If it stalls for more than a few seconds try clicking the dialogue boxes yourself.


### Generate the markdown file

The second and final step is to generate the Markdown/Commonmark file.   

When you have a directory with the same name as your .pptx file witting beside it in the current working direcotry (minus the extension) run the code with your.

    python pres2md.py  your-preso.pptx 


If you're using Pelican or another CMS that needs it you can also add a prefix to the image paths:

    python pres2md.py -i {attach} your-preso.pptx 










## Install it

This project uses pipenv. This is my first time - let me know if these instructions don't work.

First, [install pipenv](https://github.com/pypa/pipenv), eg with Brew:
    brew install pipenv

To use this:

```
git clone https://github.com/ptsefton/pptx_to_md.git`
cd pptx_to_md
pipenv shell
pipenv install
```

You _should_ end up in a virtual environment with the dependencies (python-pptx) installed and ready to go.




