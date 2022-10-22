 #!/usr/bin/env/python
import argparse
import subprocess
import os
import html
from pathlib import Path
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



def save_images_from_pdf(out_dir, pdf_filename, formatString):
    os.chdir(out_dir)
    convert_command = f"""convert  "{pdf_filename}" {formatString}"""
    print(convert_command)
    subprocess.Popen(convert_command,
                    shell=True,
                    stdout=subprocess.PIPE).communicate()




def parse_preso(powerpoint_file):
    slug = powerpoint_file.stem
    pptx = powerpoint_file.name
    pdf = powerpoint_file.with_suffix('.pdf').name
    prs = Presentation(powerpoint_file)
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
    slug: {slug}
    category:
    author:
---

<a href="{img_prefix}{pdf}">PDF version</a> | <a href="{img_prefix}{pptx}">Powerpoint Version</a>


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


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument('filename',  default=None, type=Path, help='Name of input Powerpoint file')
    parser.add_argument('-i', '--img-prefix', default = '', help="String to put in front of image paths")
    parser.add_argument('-p,', '--pdf', default = False, action="store_true", help="Convert PDF version of presentation to png images using Imagemagick convert command")
    parser.add_argument('-t,', '--topdf', default = False, action="store_true", help="Convert DOCX version to PDF using openoffice (soffice)")
    parser.add_argument('-s,', '--soffice', default = "soffice", help="soffice command")

    args = vars(parser.parse_args())

    img_prefix = args["img_prefix"]
    powerpoint_file = args["filename"].resolve()
    base_path = powerpoint_file.parent;
    pdf_file = powerpoint_file.with_suffix(".pdf")
    out_dir = base_path / powerpoint_file.stem
    md_path = out_dir / "index.md"

    if not powerpoint_file.is_file():
        raise Exception(f"{powerpoint_file} file not found")

    out_dir.mkdir(parents=True, exist_ok=True)

    shutil.copyfile(powerpoint_file, out_dir / powerpoint_file.name)
    if args["pdf"]:
        shutil.copyfile(pdf_file, out_dir / pdf_file.name)

    md, formatString = parse_preso(powerpoint_file)
    
    with open(md_path, 'w', encoding='utf-8') as o:
        print("writing", md_path)
        o.write(md)

    if args["topdf"]:
        os.chdir(out_dir)
        convert_command =  f'{args["soffice"]} --headless --convert-to pdf "{filename}"'
        print(convert_command);
        subprocess.Popen(convert_command,
                        shell=True,
                        stdout=subprocess.PIPE).communicate()
        save_images_from_pdf(out_dir, pdf_file, formatString)


    elif args["pdf"]:
        save_images_from_pdf(out_dir, pdf_file, formatString)

