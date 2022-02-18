from locale import LC_ALL
from logging import DEBUG
from typing import Optional, List
from docx.shared import Length, Inches

# Setting COMPANY_NAME which contains the name of the company
# for which the offers letters would be printed
COMPANY_NAME: str = "Celebal Technologies Private Limited"

# Path Settings
IMAGE_PATH: str = "images/"
OUTPUT_DOCX_ROOT_PATH: str = "output/document/"
OUTPUT_PDF_ROOT_PATH: str = "output/pdf/"
DOCX_TEMPLATE_PATH: str = "templates/LoiTemplate.docx"
COMPANY_SHEET_PATH: str = "CompanyInformation.xlsx"
CANDIDATE_SHEET_PATH: str = "CandidateInformation.xlsx"


# Locale Settings
LOCALE_CATEGORY: int = LC_ALL
LOCALE_TYPE: str = "en_IN.utf8"

# Log Settings
DEFAULT_LOGGER_NAME: str = "LoiProducer"
DEFAULT_LOG_MESSAGE_FORMAT: str = "%(asctime)s - %(name)s - %(levelname)s - function:%(funcName)s - line:%(lineno)d - %(message)s"
DEFAULT_LOG_LEVEL: int = DEBUG
DEFAULT_LOG_FILE_MODE: str = "w"


# Number to Word settings
DEFAULT_NUM2WORDS_LANGUAGE: str = "en_IN"

# Image Settings
COMPANY_LOGO_IMG_HEIGHT: Optional[Length] = None
COMPANY_LOGO_IMG_WIDTH: Optional[Length] = None

HR_SIGNATURE_IMG_HEIGHT: Optional[Length] = Inches(0.57)
HR_SIGNATURE_IMG_WIDTH: Optional[Length] = Inches(1.31)

CANDIDATE_SIGNATURE_IMG_HEIGHT: Optional[Length] = Inches(0.42)
CANDIDATE_SIGNATURE_IMG_WIDTH: Optional[Length] = Inches(1.09)

# File Format Settings
OUTPUT_FILE_ENDING_FORMAT: str = "_OFFER_LETTER"
ACCEPTABLE_IMAGE_FORMATS: List[str] = ["png", "jpg", "jpeg", "gif", "pjpeg", "webp", "svg"]

# Date-Time Format Settings
DEFAULT_DATE_TIME_FORMAT: str = "%d-%b-%Y"
DEFAULT_OFFER_DATE_MONTH_YEAR_FORMAT: str = " %B, %Y"
