 #!/usr/bin/env/python
import argparse
import subprocess
import os
import html
import os.path
from pptx import Presentation
import html
import math
import shutil

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
parser.add_argument('-p,', '--pdf', default = False, action="store_true", help="Convert PDF version of presentation to png images using Imagemagick convert command")
parser.add_argument('-t,', '--topdf', default = False, action="store_true", help="Convert DOCX version to PDF using openoffice (soffice)")
parser.add_argument('-s,', '--soffice', default = "soffice", help="soffice command")

args = vars(parser.parse_args())



img_prefix = args["img_prefix"]
filename = os.path.abspath(args["filename"])
path, name = os.path.split(filename)
stem, pptx_filename = os.path.splitext(name)
out_dir = os.path.join(path, stem)
md_path = os.path.join(out_dir, "index.md")
pdf_filename = stem + ".pdf"



def save_images_from_pdf(formatString):
    os.chdir(out_dir)
    convert_command = f"""convert  "{pdf_filename}" {formatString}"""
    print(convert_command)
    subprocess.Popen(convert_command,
                    shell=True,
                    stdout=subprocess.PIPE).communicate()
    



def parse_preso():
    try:
        os.mkdir(out_dir)
    except:
        pass
    shutil.copyfile(filename, os.path.join(out_dir, pptx_filename))

    prs = Presentation(filename)
    title = prs.core_properties.title
    slides = []
    for slide in prs.slides:
        s = {"text": "", "notes": ""}
        s["text"] = get_text(slide)
        if slide.has_notes_slide:
            s["notes"] = get_text(slide.notes_slide)
        slides.append(s)

    md = f"""---
    title:  >
      {title}
    date:
    slug: {stem}
    category:
    author:
---

<a href="{img_prefix}{pdf_filename}">PDF version</a> | <a href="{img_prefix}{pptx_filename}">Powerpoint Version</a>


    """


    
    digits = math.floor(math.log10(len(slides))) + 1
    count = -1
    
    formatString = "Slide%0" + str(digits) + "d.png"

    for slide in slides:
        count += 1
        # TODO make this a functions
        
        image_path = formatString % count
        md += f"""

<section typeof='http://purl.org/ontology/bibo/Slide'>
<img src='{img_prefix}{image_path}' alt='{html.escape(slide["text"], quote=True)}' title='Slide: {str(count)}' border='1'  width='85%%'/>


{slide["notes"]}


</section>

""" 
    return md, formatString

md, formatString = parse_preso()
    

with open(md_path, 'w') as o:
    print("writing", md_path)
    o.write(str(md.encode('utf-8')))

if args["topdf"]:
    os.chdir(out_dir)
    convert_command =  f'{args["soffice"]} --headless --convert-to pdf "{filename}"'
    print(convert_command);
    subprocess.Popen(convert_command,
                    shell=True,
                    stdout=subprocess.PIPE).communicate()
    save_images_from_pdf(formatString)


elif args["pdf"]:
    save_images_from_pdf(formatString)




