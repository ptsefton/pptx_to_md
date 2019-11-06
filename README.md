# pptx_to_md

Powerpoint to Markdown converter - Python 3 script



## Audience

This is for people who know how to run Python, instructions are for MacOs.

## About 

This is a simple rough and ready way to turn a Powerpoint slide deck into Markdown. It will convert the slides to PNG format (one per slide), put any text it can find on the slide into the alt attribute, and extract slide notes as a set of plain-text paragraphs.

Hint: Format your slide notes as Markdown/Commonmark for instant gratification

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

   python3 pres2md.py -p your-preso.pptx 

The first time you access a file, the operating system will make you do an authentication step.

*  If the script fails try running it again. 

*  If it stalls for more than a few seconds try clicking the dialogue boxes yourself.


### Generate the markdown file

The second and final step is to generate the Markdown/Commonmark file.   

When you have a directory with the same name as your .pptx file (minus the extension) run the code with your.

    python3 pres2md.py  your-preso.pptx 


If you're using Pelican you can also add a prefix to the image paths:

    python3 pres2md.py -i {static} your-preso.pptx 








## Install it

Make a virtual environment

Install the dependencies


