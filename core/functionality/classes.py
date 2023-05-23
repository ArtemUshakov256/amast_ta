from PIL import Image
from docxtpl import InlineImage

class ProportionalInlineImage(InlineImage):
    def __init__(self, tpl, image_descriptor, width=None, height=None):
        super().__init__(tpl, image_descriptor, width, height)

    def resize(self, max_width, max_height):
        image = Image.open(self.image_descriptor)
        image.thumbnail((max_width, max_height), Image.ANTIALIAS)
        self.width, self.height = image.size