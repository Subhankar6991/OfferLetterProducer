# from ..LoiProducer import *
import logging
import pytest
from docxtpl import RichText
import pandas as pd

import loi_producer


@pytest.mark.parametrize("a,b", [(10, 12), (15, 17)])
def test_demo(a, b):
    # context = LoiProducer.populate_company_context()
    assert a + 2 == b


def test_get_position_of_a_day():
    assert loi_producer.get_position_of_a_day(1) == "st", "Not Equal"
    assert loi_producer.get_position_of_a_day(13) == "th", "Not Equal"
    assert loi_producer.get_position_of_a_day(24) == "th", "Not Equal"
    assert loi_producer.get_position_of_a_day(0) == ""
    assert loi_producer.get_position_of_a_day(32) == ""


def test_configure_logger():
    logger = loi_producer.configure_logger("TestLogger")
    assert not logger.propagate
    assert logger.name == "TestLogger"
    assert logger.handlers[0].level == logging.DEBUG
    assert not logger.disabled
    assert logger.parent.name == "root"
    assert logger.root.name == "root"
    logger.debug("Test Logger")


@pytest.mark.parametrize(
    "file_path,file_extension",
    [
        ("/home/subhankar/Documents/test.pdf", "pdf"),
        ("/home/subhankar/Documents/test.docx", "docx"),
        ("/home/subhankar/Documents/test.doc", "doc"),
        ("/home/subhankar/Documents/test.jpg", "jpg"),
        ("/home/subhankar/Documents/test.jpeg", "jpeg"),
    ],
)
def test_get_file_extension(file_path: str, file_extension: str):
    assert loi_producer.get_file_extension(file_path) == file_extension


def test_populate_company_context(document_template, company_information, company_name):
    logger = loi_producer.configure_logger(logger_name="TestLogger", file_mode="w")
    context = loi_producer.populate_company_context(
        template=document_template,
        company_dataframe=company_information,
        company_name=company_name,
        logger_object=logger,
    )
    assert context["companyName"] == company_name
    assert context["country"] == "India"
    assert context["hrName"] == "Jonathan Smith"
    assert context["salesMail"] == "enterprisesales@celebaltech.com"
    corrupted_company_information = company_information.rename(
        columns={"companyLogo": "companyLogos"}
    )
    corrupted_context = loi_producer.populate_company_context(
        template=document_template,
        company_dataframe=corrupted_company_information,
        company_name=company_name,
        logger_object=logger,
    )
    assert corrupted_context["companyName"] == ""


def test_populate_candidate_context(document_template, candidate_information):
    logger = loi_producer.configure_logger(logger_name="TestLogger", file_mode="w")
    context = loi_producer.populate_candidate_context(
        template=document_template,
        candidate_dataframe=candidate_information,
        candidate_index=0,
        logger_object=logger,
    )
    assert context["candidateName"] == "Subhankar Karmakar"
    assert context["location"] == "Jaipur"
    assert context["designation"] == "Associate"
    assert context["basic"] == "800000"
    assert context["hra"] == "0"
    corrupted_candidate_information = candidate_information.rename(
        columns={"candidateSignature": "candidateSignatures"}
    )
    corrupted_context = loi_producer.populate_candidate_context(
        template=document_template,
        candidate_dataframe=corrupted_candidate_information,
        candidate_index=0,
        logger_object=logger,
    )
    assert corrupted_context["candidateName"] == ""


# def test_get_automapped_numeric_and_string_context():
#     assert True
#


def test_configure_rich_text_web_link(fake_document_template, fake_company_context):
    logger = loi_producer.configure_logger(logger_name="TestLogger", file_mode="w")
    rt = loi_producer.configure_rich_text_web_link(
        template=fake_document_template,
        company_dataframe=pd.DataFrame([fake_company_context]).set_index("companyName"),
        company_name="Test Company Name",
        logger_object=logger,
    )

    assert type(rt) == RichText


def test_configure_rich_text_date_of_offer(
    fake_document_template, fake_candidate_context_list
):
    logger = loi_producer.configure_logger(logger_name="TestLogger", file_mode="w")
    df = pd.DataFrame(fake_candidate_context_list)
    df["offerDate"] = pd.to_datetime(df["offerDate"])
    rt = loi_producer.configure_rich_text_date_of_offer(
        candidate_dataframe=df, candidate_index=0, logger_object=logger
    )

    assert type(rt) == RichText


def test_render_and_produce_PDF(
    fake_document_template, fake_company_context, fake_candidate_context, mocker
):
    mocker.patch(
        "loi_producer.OUTPUT_DOCX_ROOT_PATH", "tests/test_output/test_document/"
    )
    mocker.patch("loi_producer.OUTPUT_PDF_ROOT_PATH", "tests/test_output/test_pdf/")
    mocker.patch("loi_producer.OUTPUT_FILE_ENDING_FORMAT", "_test_")
    # mocker.patch("loi_producer.DOCX_TEMPLATE_PATH", "tests/tes_templates/test_loi_template.docx")
    # mocker.patch("conftest.loi_producer_config.DOCX_TEMPLATE_PATH", "tests/test_templates/test_loi_template.docx")

    # mocker.patch("loi_producer.populate_candidate_context", return_value=fake_candidate_context)
    # mocker.patch("loi_producer.populate_company_context", return_value=fake_company_context)

    context = {
        **fake_candidate_context,
        **fake_company_context,
    }
    logger = loi_producer.configure_logger(logger_name="TestLogger", file_mode="w")
    loi_producer.render_and_produce_PDF(
        template=fake_document_template,
        context_information=context,
        candidate_name=fake_candidate_context["candidateName"] or "CORRUPTED",
        logger_object=logger,
    )


def test_main(
    fake_document_template, fake_company_context, fake_candidate_context_list, mocker
):
    mocker.patch(
        "loi_producer.DOCX_TEMPLATE_PATH", "tests/test_templates/test_loi_template.docx"
    )
    mocker.patch("loi_producer.OUTPUT_FILE_ENDING_FORMAT", "_test_main")
    mocker.patch(
        "loi_producer.pd.read_excel",
        side_effect=[fake_candidate_context_list, fake_company_context],
    )
    mocker.patch(
        "loi_producer.populate_company_context", return_value=fake_company_context
    )
    mocker.patch(
        "loi_producer.configure_rich_text_web_link", return_value="fake_web_link"
    )
    mocker.patch(
        "loi_producer.configure_rich_text_date_of_offer", return_value="fake_offer_date"
    )
    mocker.patch(
        "loi_producer.populate_candidate_context",
        side_effect=fake_candidate_context_list,
    )

    # mocker.patch("loi_producer.OUTPUT_DOCX_ROOT_PATH", "tests/test_output/test_document/")
    # mocker.patch("loi_producer.OUTPUT_PDF_ROOT_PATH", "tests/test_output/test_pdf/")
    mocker.patch("loi_producer.render_and_produce_PDF", return_value=None)

    logger = loi_producer.configure_logger(logger_name="TestLogger", file_mode="w")
    loi_producer.main(company_name="TestCompany", logger_object=logger)
