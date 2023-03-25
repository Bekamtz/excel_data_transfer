import pandas as pd
from openpyxl import load_workbook
from datetime import datetime
from data_transfer import DataTransfer
import logging.config
import yaml

with open('logging_conf.yaml', 'rt') as f:
    config = yaml.safe_load(f.read())
logging.config.dictConfig(config)
logger = logging.getLogger('data_transfer')

INPUT_FILE = 'input_file.xlsx'
TARGET_FILE = 'target_file.xlsx'
DATE = '2023/07/31'

def run():
    try:
        logger.info('Data Transfer Started')
    # Instantiate Class 
        logger.info('Intantiating Class')
        dt = DataTransfer(input_file=INPUT_FILE, target_file=TARGET_FILE, date=DATE)
        logger.info('Starting Data Transfer')
    # Long Strategy Exposure
        dt.transfer_data(input_names_row_range=(8, 14), input_values_row_range=(8, 14), target_names_row_range=(504, 510))

    # Long Credit Exposure
        dt.transfer_data(input_names_row_range=(19, 36), input_values_row_range=(19, 36), target_names_row_range=(512, 526))

    # Long Geographic Exposure
        dt.transfer_data(input_names_row_range=(8, 14), input_values_row_range=(8, 14), target_names_row_range=(535, 541), input_values_col=9,input_names_col=8)

    # Long Industry Sector Exposure
        dt.transfer_data(input_names_row_range=(18, 36), input_values_row_range=(18, 36), target_names_row_range=(541, 554), input_values_col=9,input_names_col=8)

    # Long Market Exposure
        dt.transfer_data(input_names_row_range=(37, 41), input_values_row_range=(37, 41), target_names_row_range=(555, 559), input_values_col=9,input_names_col=8)

    # Long Sovereign Exposure
        dt.transfer_data(input_names_row_range=(45, 52), input_values_row_range=(45, 52), target_names_row_range=(560, 564), input_values_col=9,input_names_col=8)

    # Short Strategy Exposure
        dt.transfer_data(input_names_row_range=(8, 14), input_values_row_range=(8, 14), target_names_row_range=(569, 575),input_values_col=2)

    # Short Credit Exposure
        dt.transfer_data(input_names_row_range=(19, 36), input_values_row_range=(19, 36), target_names_row_range=(577, 591),input_values_col=2)

    # Short Geographic Exposure
        dt.transfer_data(input_names_row_range=(8, 14), input_values_row_range=(8, 14), target_names_row_range=(601, 605), input_values_col=10,input_names_col=8)

    # Short Industry Sector Exposure
        dt.transfer_data(input_names_row_range=(18, 36), input_values_row_range=(18, 36), target_names_row_range=(606, 619), input_values_col=10,input_names_col=8)

    # Short Market Exposure
        dt.transfer_data(input_names_row_range=(37, 41), input_values_row_range=(37, 41), target_names_row_range=(620, 624), input_values_col=10,input_names_col=8)

    # Short Sovereign Exposure
        dt.transfer_data(input_names_row_range=(45, 52), input_values_row_range=(45, 52), target_names_row_range=(625, 629), input_values_col=10,input_names_col=8)

    # Long Total Privates
        dt.transfer_data(input_names_row_range = (64,65),input_values_row_range=(64,65),target_names_row_range=(532,533))

    #Short Total Privates
        dt.transfer_data(input_names_row_range = (597,598),input_values_row_range=(597,598),target_names_row_range=(532,533),input_values_col=2)

    #  Long Unadjusted Portfolio
        dt.transfer_data(input_names_row_range = (66,68),input_values_row_range=(66,68),target_names_row_range=(533,534))

    #Short Unadjusted Porfolio
        dt.transfer_data(input_names_row_range = (598,599),input_values_row_range=(598,599),target_names_row_range=(533,534),input_values_col=2)
        logger.info('Data Transfer Complete')
    except:
        logger.error('Uh-Oh Data Transfer Failed')

if __name__ == "__main__":
    run()

