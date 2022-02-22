# from ..LoiProducer import *
import pytest
from docxtpl import DocxTemplate
import loi_producer
import pandas as pd


@pytest.mark.parametrize("a,b", [(10, 12), (15, 17)])
def test_demo(a, b):
    # context = LoiProducer.populate_company_context()
    assert a + 2 == b


def test_get_position_of_a_day():
    assert str(1) + loi_producer.get_position_of_a_day(1) == "1st", "Not Equal"
    assert str(13) + loi_producer.get_position_of_a_day(13) == "13th", "Not Equal"
    assert str(24) + loi_producer.get_position_of_a_day(24) == "24th", "Not Equal"


def test_company_context(document_template, company_information, company_name):
    context = loi_producer.populate_company_context(
        template=document_template,
        company_dataframe=company_information,
        company_name=company_name,
    )
    assert context["companyName"] == company_name
    assert context["country"] == "India"
    assert context["hrName"] == "Jonathan Smith"
    assert context["salesMail"] == "enterprisesales@celebaltech.com"


def test_candidate_context(document_template, candidate_information):
    context = loi_producer.populate_candidate_context(
        template=document_template,
        candidate_dataframe=candidate_information,
        candidate_index=0,
    )
    assert context["candidateName"] == "Subhankar Karmakar"
    assert context["location"] == "Jaipur"
    assert context["designation"] == "Associate"
    assert context["basic"] == "800000"
    assert context["hra"] == "0"
