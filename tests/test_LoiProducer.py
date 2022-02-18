# from ..LoiProducer import *
import pytest
from docxtpl import DocxTemplate
import LoiProducer
import pandas as pd

COMPANY_NAME: str = ""


@pytest.fixture(autouse=True)
def company_name():
    global COMPANY_NAME
    COMPANY_NAME = "Celebal Technologies Private Limited"


@pytest.fixture()
def company_information():
    return pd.read_excel(
        r"C:\Users\User\Desktop\Celebal Intern\task3\CompanyInformation.xlsx",
        index_col="companyName",
    )


@pytest.fixture()
def candidate_information():
    return pd.read_excel(
        r"C:\Users\User\Desktop\Celebal Intern\task3\CandidateInformation.xlsx",
    )


@pytest.fixture()
def document_template():
    return DocxTemplate(
        r"C:\Users\User\Desktop\Celebal Intern\task3\templates\LoiTemplate.docx"
    )


@pytest.mark.parametrize("a,b", [(10, 12), (15, 17)])
def test_demo(a, b):
    # context = LoiProducer.populate_company_context()
    assert a + 2 == b


def test_get_position_of_a_day():
    assert str(1) + LoiProducer.get_position_of_a_day(1) == "1st", "Not Equal"
    assert str(13) + LoiProducer.get_position_of_a_day(13) == "13th", "Not Equal"
    assert str(24) + LoiProducer.get_position_of_a_day(24) == "24th", "Not Equal"


# @pytest.mark.usefixtures("row_identifier")
def test_company_context(document_template, company_information, company_name):
    context = LoiProducer.populate_company_context(
        template=document_template, company_dataframe=company_information, company_name=COMPANY_NAME
    )
    assert context["companyName"] == COMPANY_NAME
    assert context["country"] == "India"
    assert context["hrName"] == "Jonathan Smith"
    assert context["salesMail"] == "enterprisesales@celebaltech.com"


def test_candidate_context(document_template, candidate_information):
    context = LoiProducer.populate_candidate_context(
        template=document_template, candidate_dataframe=candidate_information, candidate_index=0
    )
    assert context["candidateName"] == "Subhankar Karmakar"
    assert context["location"] == "Jaipur"
    assert context["designation"] == "Associate"
    assert context["basic"] == "800000"
    assert context["hra"] == "0"
