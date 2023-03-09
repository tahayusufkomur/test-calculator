import os


def create_dirs(path):
    if os.path.isdir(f"{path}/raporlar"):
        import shutil

        shutil.rmtree(f"{path}/raporlar")

    os.mkdir(f"{path}/raporlar")
    os.mkdir(f"{path}/raporlar/text_reports")
    os.mkdir(f"{path}/raporlar/excel_reports")
    os.mkdir(f"{path}/raporlar/temp_files")
    os.mkdir(f"{path}/raporlar/text_reports/person_reports")
    os.mkdir(f"{path}/raporlar/excel_reports/mlq_reports")


def common_paths(path):
    return \
        f"{path}/../resources/yanıtlar/", \
        f"{path}/raporlar/excel_reports", \
        f"{path}/raporlar/text_reports", \
        f"{path}/raporlar/temp_files", \
        f"{path}/../resources/passwords.json"
