import os
import pathlib
import zipfile

from django.contrib.auth.decorators import login_required
from django.core.files.storage import FileSystemStorage
from django.http import HttpResponse
from django.shortcuts import render

CURRENT_DIR = pathlib.Path(__file__).parent.resolve()


@login_required
def index(response):
    return render(response, 'base.html', {})


@login_required
def home(response):
    return render(response, "home.html", {})


@login_required
def test_calculator(request):
    if request.method == 'POST' and request.FILES['myfile']:

        # get file
        myfile = request.FILES['myfile']
        fs = FileSystemStorage()
        fs.save(myfile.name, myfile)

        # unzip to proper path
        with zipfile.ZipFile(myfile.name, 'r') as zip_ref:
            zip_ref.extractall(f"{CURRENT_DIR}/../src/")

        # process
        from src.main import main
        main()

        # zip
        import shutil
        shutil.make_archive("reports", 'zip', f"{CURRENT_DIR}/../src/raporlar")

        # return
        zip_file = open("reports.zip", 'rb')
        response = HttpResponse(zip_file, content_type='application/force-download')
        response['Content-Disposition'] = 'attachment; filename="%s"' % 'reports.zip'

        # clean generated files
        os.remove("files.zip")
        os.remove("reports.zip")
        src_dir = os.listdir(f"{CURRENT_DIR}/../src")
        valid_files = ['__init__.py', 'passwords', 'personality_tests', 'base_personality_test.py', 'main.py', 'utilities.py']
        for file in src_dir:
            if file not in valid_files:
                file_path = f"{CURRENT_DIR}/../src/{file}"
                if os.path.isdir(file_path):
                    print(f"dir removed {file_path}")
                    shutil.rmtree(file_path)
                else:
                    os.remove(file_path)
                    print(f"file removed {file_path}")
        return response

    return render(request, 'test_calculator/index.html')
