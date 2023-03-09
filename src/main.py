import json
import pathlib
import string
import pandas as pd
from src.personality_tests.attendance import Attendance
from src.personality_tests.cipto import Cipto
from src.personality_tests.cmvkb import Cmvkb
from src.personality_tests.minesota import Minesota
from src.personality_tests.mlq_ast import MlqAst
from src.personality_tests.mlq_schema import MlqSchema
from src.personality_tests.performance import Performance
from src.personality_tests.plo import PLO
from src.personality_tests.rotterdam import Rotterdam
from src.personality_tests.varis_performance import VarisPerformance
from src.utilities import create_text_report
from src.personality_tests.five_factor import FiveFactor

CURRENT_DIR = pathlib.Path(__file__).parent.resolve()
color_file_path = f"{CURRENT_DIR}/files/colors.json"
answers_path = f"{CURRENT_DIR}/files/excel_answers"
passwords_path = f"{CURRENT_DIR}/files/passwords.json"


def main():
    if color_file_path:
        with open(color_file_path) as json_file:
            colors = json.load(json_file)

        for key in colors:
            if colors[key] == "":
                colors[key] = None

    import os

    file_names = os.listdir(answers_path)
    files = {}

    for file in file_names:
        if file.endswith("xlsx"):
            if not file.startswith('~'):
                file_path = f"{answers_path}/{file}"
                file_name = os.path.splitext(file)[0]
                files[file_name] = file_path

    if os.path.isdir(f"{CURRENT_DIR}/raporlar"):
        import shutil
        shutil.rmtree(f"{CURRENT_DIR}/raporlar")

    os.mkdir(f"{CURRENT_DIR}/raporlar")
    os.mkdir(f"{CURRENT_DIR}/raporlar/text_reports")
    os.mkdir(f"{CURRENT_DIR}/raporlar/excel_reports")
    os.mkdir(f"{CURRENT_DIR}/raporlar/temp_files")
    os.mkdir(f"{CURRENT_DIR}/raporlar/text_reports/person_reports")
    os.mkdir(f"{CURRENT_DIR}/raporlar/excel_reports/mlq_reports")

    # create a xlsx file to join texts based on names
    pws = json.load(open(passwords_path, encoding="utf-8"))
    if 'regex' in pws:
        del pws['regex']
    df = pd.DataFrame(columns=['İsim Soyisim'])
    pws_values = [string.capwords(i) for i in pws.values()]
    df['İsim Soyisim'] = pws_values
    df.index = df['İsim Soyisim']
    # df.drop(['İsim Soyisim'], axis=1, inplace=True)
    df.to_excel(f"{CURRENT_DIR}/raporlar/temp_files/all_reports.xlsx", engine='openpyxl', index=False)

    # five factor test
    if "b5kt" in files:
        five_factor_test = FiveFactor(files["b5kt"], colors=colors)
        five_factor_test.create_report()

    # rotterdam
    if "rdo" in files:  # rotterdam
        rotterdam_test = Rotterdam(files["rdo"], colors=colors)
        rotterdam_test.create_report()

    # PLO potansiyel liderlik ölçeği
    if "plo" in files:  # PLO
        plo_test = PLO(files['plo'], colors=colors)
        plo_test.create_report()

    # # # ÇİPTÖ: Çalışanların iş yerindeki problemlere verdikleri tepkiler ölçeği
    if "cipto" in files:  #
        cipto_test = Cipto(files['cipto'], colors=colors)
        cipto_test.create_report()

    # minesota
    if "mito" in files:  #
        minesota_test = Minesota(files['mito'], colors=colors)
        minesota_test.create_report()

    # Çalışan motivasyonu ve kurumsal bağlılık testi
    if "cmvkb" in files:  #
        cmvkb_test = Cmvkb(files['cmvkb'], colors=colors)
        cmvkb_test.create_report()

    # MLQ Çoklu Liderlik Alt Faktörleri
    if "mlq-ast" and "mlq-ust" in files:
        mlqast_test = MlqAst(non_manager_path=files['mlq-ast'], manager_path=files['mlq-ust'], colors=colors)
        mlqast_test.create_report()

        # create schema report
        MlqSchema(non_manager_path=files['mlq-ast'], manager_path=files['mlq-ust'], colors=colors).create_report()

    if "performance" in files:  #
        performance_test = Performance(files['performance'], colors=colors)
        performance_test.create_report()

    if "Varis" in files:  #
        performance_test = VarisPerformance(files['Varis'], colors=colors)
        performance_test.create_report()

    if files:
        Attendance(files).create_report()
    #
    try:
        create_text_report(f"{CURRENT_DIR}/raporlar/temp_files", f"{CURRENT_DIR}/raporlar/text_reports/person_reports")
    except Exception as error:
        print(error)

    # remove temp files
    if os.path.isdir(f'{CURRENT_DIR}/raporlar/temp_files'):
        import shutil
        shutil.rmtree(f'{CURRENT_DIR}/raporlar/temp_files')


if __name__ == '__main__':
    main()