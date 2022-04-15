########################################################################################################################
# Filename: shape_utils.py
#
# Purpose: Utilities for pptx.shapes.
# Author(s): Bobby (Robert) Lumpkin
#
# Library Dependencies: 
########################################################################################################################

import pandas as pd
import numpy as np
import io
import pptx
from pptx.util import Inches, Pt

def init_pres(
    title: str,
    template: str = None,
    save: bool = False,
    save_path: str = None
):
    """
    Initializes a pptx.Presentation object.
    """
    ppt = pptx.Presentation()
    if template:
        ppt = pptx.Presentation(template)
    slide = ppt.slides.add_slide(ppt.slide_layouts[0])
    slide.shapes.title.text = title

    # Save, if appropriate
    if save:
        ppt.save(save_path)

    return ppt