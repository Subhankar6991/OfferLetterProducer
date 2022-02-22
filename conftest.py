import os
import sys

import pandas as pd
import pytest
from docxtpl import DocxTemplate

import loi_producer_config as config

myPath = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, myPath)


@pytest.fixture(autouse=True)
def company_name():
    return "Celebal Technologies Private Limited"


@pytest.fixture()
def company_information():
    return pd.read_excel(
        config.COMPANY_SHEET_PATH,
        index_col="companyName",
    )


@pytest.fixture()
def candidate_information():
    return pd.read_excel(
        config.CANDIDATE_SHEET_PATH,
    )


@pytest.fixture()
def document_template():
    return DocxTemplate(
        config.DOCX_TEMPLATE_PATH
    )
