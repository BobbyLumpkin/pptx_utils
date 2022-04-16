########################################################################################################################
# Filename: test_shape_utils.py
#
# Purpose: Tests for shapes_utils.
# Author(s): Bobby (Robert) Lumpkin
#
# Library Dependencies: 
########################################################################################################################

import pytest
import pptx
import pandas as pd
import io
import sys

sys.path.append("C:\\Users\\rober\\OneDrive\\Documents\\Github_Repos\\pptx_utils\\src")
from shape_utils import init_pres, table_from_df, matplotlib_to_pic


@pytest.mark.shape_utils
def test_init_pres():
    """
    Tests the functionality of 'init_pres()'.
    """
    ppt_stream = io.BytesIO()
    # Call init_pres()
    ppt = init_pres(
        title = "Test Presentation",
        template = None,
        save = True,
        save_path = ppt_stream
    )
    
    # Check save functionality
    assert ppt_stream.getbuffer().nbytes > 0
    loaded_ppt = pptx.Presentation(ppt_stream)

    # Check slide title
    try:
        assert ppt.slides[0].shapes.title.text == "Test Presentation"
        assert loaded_ppt.slides[0].shapes.title.text == "Test Presentation"
    except IndexError:
        assert False


#@pytest.mark.shape_utils
def test_table_from_df(

):
    """
    Tets functionality of 'table_from_df()'.
    """

    return
    