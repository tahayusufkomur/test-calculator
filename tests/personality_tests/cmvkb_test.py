import pytest
from src.personality_tests.cmvkb import Cmvkb
import pandas as pd
import numpy as np
import os
from src.utilities import create_dirs, common_paths
import pathlib

# get current path
path = pathlib.Path(__file__).parent.resolve()

@pytest.fixture
def cmvkb():
    # init directories
    if os.path.isdir('raporlar'):
        import shutil
        shutil.rmtree('raporlar')


    # create necessary folders
    create_dirs(path)

    # get paths
    test_path, output_path, text_output_dir, temp_dir, password_path = common_paths(path)

    return Cmvkb(test_path=test_path+"ÇMVKB.xlsx",
                 output_path=output_path,
                 text_output_path=text_output_dir,
                 temp_dir=temp_dir,
                 passwords_path=password_path
                 )


def test_cmvkb(cmvkb):
    cmvkb.create_report()

    expected_df_path = f"{path}/../resources/expected_reports/ÇMVKB.xlsx"

    result_df = pd.read_excel(cmvkb.output_path, engine='openpyxl')
    result_df = result_df.dropna()

    # convert float64 columns to int64
    for column in result_df.columns:
        if result_df[column].dtype == np.float64:
            result_df[column] = result_df[column].astype(np.int64)

    expected_df = pd.read_excel(expected_df_path, engine='openpyxl')
    expected_df.dropna(inplace=True)

    for index, row in expected_df.iterrows():
        expected = expected_df.iloc[[index]]
        result = result_df.iloc[[index]]
        assert result.equals(expected)
