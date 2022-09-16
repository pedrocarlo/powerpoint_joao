import os
import pptx.shapes.picture
from pptx import Presentation
from pptx.util import Inches
from typing import List

BLANK_SLIDE = 6
slide_height = Inches(7.5).emu  # Default Slide Height
slide_width = Inches(13.34).emu  # Default Slide Width


class Slide:
    def __init__(self, template_file: str, images: List[str] = None, image_order: List[int] = None,
                 texts: List[str] = None):
        if images is not None:
            self.images_with_correct_index = sorted([x for x in zip(images, image_order)], key=lambda x: x[1])
        else:
            self.images_with_correct_index = None
        self.texts = texts  # stores the texts for the textboxes
        self.template_file = template_file
        self.presentation = Presentation(template_file)
        # read template and determine number of pictures and textboxes
        self.num_images, self.num_text = self.number_pictures_text()

    def update_images(self, images: List[str], image_order: List[int]):
        self.images_with_correct_index = sorted([x for x in zip(images, image_order)], key=lambda x: x[1])

    def number_pictures_text(self):
        img_count = 0
        text_count = 0
        for slide in self.presentation.slides:
            for shape in slide.shapes:
                if isinstance(shape,
                              pptx.shapes.autoshape.Shape) and shape.has_text_frame:  # In this case the autoshape in question is a textbox
                    text_count += 1
                if isinstance(shape, pptx.shapes.picture.Picture):  # If object is a Picture
                    img_count += 1
        return img_count, text_count

    def add_slide(self, presentation: Presentation):
        add_pictures_text(presentation, Presentation(self.template_file), self)


def new_presentation(title):
    new_prs = Presentation()
    return new_prs, title


prs = Presentation("templates/test2.pptx")
prs.slide_height = slide_height
prs.slide_width = slide_width
images = os.listdir("imagens")


def add_pictures_text(presentation: Presentation, template: Presentation, new_slide: Slide):
    presentation.slide_height = template.slide_height
    presentation.slide_width = template.slide_width
    new_pres_slide = presentation.slides.add_slide(blank_slide_layout)
    img_count = 0
    text_count = 0
    for slide in template.slides:
        for shape in slide.shapes:
            if isinstance(shape,
                          pptx.shapes.autoshape.Shape) and shape.has_text_frame:  # In this case the autoshape in
                # question is a textbox
                try:
                    left = shape.left
                    top = shape.top
                    height = shape.height
                    width = shape.width
                    txBox = new_pres_slide.shapes.add_textbox(left, top, width, height)
                    tf = txBox.text_frame
                    tf.text = new_slide.texts[text_count]  # TODO CHANGE TEXT HERE TO BE FROM SLIDE
                except IndexError:  # If user does not supply enough text boxes the code will still run
                    continue
                text_count += 1
            if isinstance(shape, pptx.shapes.picture.Picture):  # If object is a Picture
                try:
                    left = shape.left
                    top = shape.top
                    height = shape.height
                    # print(images)
                    # print(img_count)
                    new_pres_slide.shapes.add_picture(new_slide.images_with_correct_index[img_count][0],
                                                      left, top, height=height)  # TODO get images from slide
                except (IndexError, FileNotFoundError,
                        FileExistsError):  # if user does not submit enough photos it will stil run
                    continue
                img_count += 1


# new_prs, title = new_presentation("test3")
# blank_slide_layout = new_prs.slide_layouts[6]
# add_pictures_text(new_prs, prs, None)
# new_prs.save(title + '.pptx')
