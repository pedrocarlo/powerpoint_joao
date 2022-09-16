import pptx.shapes.picture, os
from pptx import Presentation
from pptx.util import Inches, Pt, Centipoints

BLANK_SLIDE = 6
slide_height = Inches(7.5).emu  # Default Slide Height
slide_width = Inches(13.34).emu  # Default Slide Width


def new_presentation(title):
    new_prs = Presentation()
    new_prs.slide_height = slide_height
    new_prs.slide_width = slide_width


# this code only works for templates with one slide
def number_pictures_text(presentation: Presentation):
    img_count = 0
    text_count = 0
    for slide in presentation.slides:
        for shape in slide.shapes:
            if isinstance(shape,
                          pptx.shapes.autoshape.Shape) and shape.has_text_frame:  # In this case the autoshape in question is a textbox
                text_count += 1
            if isinstance(shape, pptx.shapes.picture.Picture):  # If object is a Picture
                img_count += 1
    return img_count, text_count


prs = Presentation("templates/test2.pptx")
new_prs = Presentation()
new_prs.slide_height = slide_height  # Default Slide Height
new_prs.slide_width = slide_width  # Default Slide Width

blank_slide_layout = new_prs.slide_layouts[6]
images = os.listdir("imagens")

def add_pictures_text(presentation):


for slide in prs.slides:
    img_count = 0
    text_count = 0
    new_slide = new_prs.slides.add_slide(blank_slide_layout)  # create a new blank slide in new presentation
    for shape in slide.shapes:
        if isinstance(shape,
                      pptx.shapes.autoshape.Shape) and shape.has_text_frame:  # In this case the autoshape in question is a textbox
            left = shape.left
            top = shape.top
            height = shape.height
            width = shape.width
            txBox = new_slide.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            tf.text = shape.text
            text_count += 1
        if isinstance(shape, pptx.shapes.picture.Picture):  # If object is a Picture
            left = shape.left
            top = shape.top
            height = shape.height
            # print(images)
            # print(img_count)
            new_slide.shapes.add_picture("imagens/" + images[img_count], left, top, height=height)
            img_count += 1

new_prs.save('test3.pptx')
