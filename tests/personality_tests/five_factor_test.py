import os
import pytest
from src.personality_tests.five_factor import FiveFactor
import pandas as pd
import numpy as np
from src.utilities import create_dirs, common_paths
import pathlib

# get current path
path = pathlib.Path(__file__).parent.resolve()


@pytest.fixture
def five_factor():
    # init directories
    if os.path.isdir('raporlar'):
        import shutil
        shutil.rmtree('raporlar')


    # create necessary folders
    create_dirs(path)

    # get paths
    test_path, output_path, text_output_dir, temp_dir, password_path = common_paths(path)

    return FiveFactor(test_path=test_path + "B5KT.xlsx",
                      output_path=output_path,
                      text_output_path=text_output_dir,
                      temp_dir=temp_dir,
                      passwords_path=password_path
                      )


def test_five_factor(five_factor):
    five_factor.create_report()

    expected_df_path = f"{path}/../resources/expected_reports/B5KT.xlsx"

    result_df = pd.read_excel(five_factor.output_path, engine='openpyxl')
    result_df = result_df.dropna()

    # convert float64 columns to int64
    for column in result_df.columns:
        if result_df[column].dtype == np.float64:
            result_df[column] = result_df[column].astype(np.int64)

    expected_df = pd.read_excel(expected_df_path, engine='openpyxl')
    expected_df.dropna(inplace=True)

    for column in expected_df.columns:
        if expected_df[column].dtype == np.float64:
            expected_df[column] = expected_df[column].astype(np.int64)

    # result_df.to_excel('xxx.xlsx', engine='openpyxl')
    for index, row in result_df.iterrows():
        for col in result_df.columns:
            result = expected_df.iloc[[index]][col]
            expected = result_df.iloc[[index]][col]
            assert result.equals(expected), '\n{0}: {1} --- {2} != {3}'.format(row['İsim Soyisim'], col, result,
                                                                               expected)
