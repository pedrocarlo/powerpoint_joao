from typing import List

from pptx import Presentation

from templates import number_pictures_text


class Slide:
    def __init__(self, template_file: str, images: List[str], image_order: List[int]):
        self.images_with_correct_index = sorted([x for x in zip(images, image_order)], key=lambda x: x[1])
        self.template_file = template_file
        self.presentation = Presentation(template_file)
        # read template and determine number of pictures and textboxes
        self.num_images, self.num_text = number_pictures_text(self.presentation)

    def add