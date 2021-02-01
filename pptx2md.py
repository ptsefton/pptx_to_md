 #!/usr/bin/env/python
import argparse
import subprocess
import os
import html
import parser
import os.path
from pptx import Presentation
import html
import math

def get_text(slide):
    text = ""
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        for paragraph in shape.text_frame.paragraphs:
            for run in paragraph.runs:
                text += run.text
            text += "\n"
    return(text)


parser = argparse.ArgumentParser()
parser.add_argument('filename',  default=None, help='Name of input Powerpoint file')
parser.add_argument('-i', '--img-prefix', default = '', help="String to put in front of image paths")
parser.add_argument('-p,', '--powerpoint', default = False, action="store_true", help="Attempt to presentation to png images using Powerpoint first")

args = vars(parser.parse_args())
import glob
import shutil


img_prefix = args["img_prefix"]
filename = os.path.abspath(args["filename"])
path, name = os.path.split(filename)
stem, _ = os.path.splitext(name)
out_dir = os.path.join(path, stem)

md_path = os.path.join(out_dir, "index.md")


def save_images_using_powerpoint():

    ppt_command = 'osascript pptx_to_png_powerpoint.scrpt "%s"  "%s" "%s"' % (out_dir, filename, stem)
    
    subprocess.Popen(ppt_command,
                               shell=True,
                               stdout=subprocess.PIPE).communicate()




def parse_preso():
    prs = Presentation(filename)
    title = prs.core_properties.title
    slides = []
    for slide in prs.slides:
        s = {"text": "", "notes": ""}
        s["text"] = get_text(slide)
        if slide.has_notes_slide:
            s["notes"] = get_text(slide.notes_slide)
        slides.append(s)

    md = """---
    title:  >
      %s
    date:
    slug: %s
    category:
    author:
---



    """ % (title, stem)


    count = 0
    digits = math.floor(math.log10(len(slides))) + 1
    formatString = "Slide%0" + str(digits) + "d.png"
    for slide in slides:
        count += 1
        # TODO make this a functions
        
        image_path = formatString % count
        md += """
<section typeof='http://purl.org/ontology/bibo/Slide'>
<img src='%s%s' alt='%s' title='%s' border='1'  width='85%%'/>


%s


</section>

"""  % (img_prefix, image_path, html.escape(slide["text"], quote=True), str(count), slide["notes"])
    return md


if args["powerpoint"]:
    save_images_using_powerpoint()
else:
    md = parse_preso()
    

with open(md_path, 'w') as o:
    print("writing", md_path)
    o.write(md)
