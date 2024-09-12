import os
import pandas as pd
import numpy as np
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Emu, Pt
from tqdm import tqdm


class KidaseCreator:
    '''
    This class is responsible for aggregating the kidase data and components 
    to create the slides for the Liturgy, readings, and hymns.
    '''
    def __init__(self, data_path, languages):
        self.order_of_the_liturgy = pd.read_excel(os.path.join(data_path, 'ሥርዓተ_ቅዳሴ.xlsx'))
        self.languages = languages
        self.slide_settings = {
            'background_color': RGBColor(0, 0, 0),
            'font': 'Arial',
            'font_size': Pt(24),
            'font_color': RGBColor(255, 255, 255),  
            'border_color': RGBColor(255, 255, 255),
            'border_thickness': Pt(1)  
        }
        self.process_data()

    def process_data(self):
        '''
        Checks the order of the liturgy for the given languages.
        '''
        for lang in self.languages:
            if lang not in self.order_of_the_liturgy.columns:
                raise ValueError(f"Language '{lang}' not found in the data.")
            self.order_of_the_liturgy[lang] = self.order_of_the_liturgy[lang].fillna('')

    def format_slide(self, slide, slide_settings):
        '''
        Prepare the slide for the given languages.
        '''
        left, top, width, height = slide_settings['text_box_position']
        textbox = slide.shapes.add_textbox(left, top, width, height)
        p = textbox.text_frame.add_paragraph()
        p.text = slide_settings['text']
        p.font.name = slide_settings['font']
        p.font.size = slide_settings['font_size']
        p.font.color.rgb = slide_settings['font_color']
        textbox.line.color.rgb = slide_settings['border_color']
        textbox.line.width = slide_settings['border_thickness']

    def create_presentation(self):
        # Create a presentation object
        prs = Presentation()
        prs.slide_height = Emu(5.5 * 914400) # 5.5 inches to EMU
        slide_layout = prs.slide_layouts[6]  # Blank layout
        for i in tqdm(range(len(self.order_of_the_liturgy)), desc="Creating slides"):
            slide = prs.slides.add_slide(slide_layout)
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = self.slide_settings['background_color']  # Black color
            if len(self.languages) in [1, 2, 3]:
                # Use columns for up to 3 languages
                width = prs.slide_width // len(self.languages)
                for j, lang in enumerate(self.languages):
                    self.slide_settings['text_box_position'] = (j * width, 0, width, prs.slide_height)
                    self.slide_settings['text'] = self.order_of_the_liturgy[lang][i]
                    self.format_slide(slide, self.slide_settings)
            elif len(self.languages) == 4:
                # Use 2x2 grid for 4 languages
                width = prs.slide_width // 2
                height = prs.slide_height // 2
                for j, lang in enumerate(self.languages):
                    self.slide_settings['text_box_position'] = ((j % 2) * width, (j // 2) * height, width, height)
                    self.slide_settings['text'] = self.order_of_the_liturgy[lang][i]
                    self.format_slide(slide, self.slide_settings)
            else:
                raise ValueError("Only up to 4 languages are supported.")
        # Save the presentation
        prs.save('test.pptx')
