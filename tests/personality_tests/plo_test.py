import os

import pytest
from src.personality_tests.plo import PLO
import pandas as pd
import numpy as np
from src.utilities import create_dirs, common_paths
import pathlib

# get current path
path = pathlib.Path(__file__).parent.resolve()

@pytest.fixture
def plo():
    # init directories
    if os.path.isdir('raporlar'):
        import shutil
        shutil.rmtree('raporlar')

    # create necessary folders
    create_dirs(path)

    # get paths
    test_path, output_path, text_output_dir, temp_dir, password_path = common_paths(path)

    return PLO(test_path=test_path+"PLO.xlsx",
               output_path=output_path,
               text_output_path=text_output_dir,
               temp_dir=temp_dir,
               passwords_path=password_path
               )


def test_plo(plo):
    plo.create_report()

    expected_df_path = f"{path}/../resources/expected_reports/PLO.xlsx"

    result_df = pd.read_excel(plo.output_path, engine='openpyxl')
    result_df = result_df.dropna()

    # convert float64 columns to int64
    for column in result_df.columns:
        if result_df[column].dtype == np.float64:
            result_df[column] = result_df[column].astype(np.int64)

    expected_df = pd.read_excel(expected_df_path, engine='openpyxl')

    for index, row in expected_df.iterrows():
        expected = expected_df.iloc[[index]]
        result = result_df.iloc[[index]]
        assert result.equals(expected)
