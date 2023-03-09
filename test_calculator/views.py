import os
import pathlib
import zipfile
from django.core.files.storage import FileSystemStorage
from django.http import HttpResponse
from django.shortcuts import render

CURRENT_DIR = pathlib.Path(__file__).parent.resolve()


# @login_required
def index(response):
    return render(response, 'base.html', {})


# @login_required
def home(response):
    return render(response, "home.html", {})


def test_calculator(request):
    if request.method == 'POST' and request.FILES['myfile']:

        # get file
        myfile = request.FILES['myfile']
        fs = FileSystemStorage()
        fs.save(myfile.name, myfile)

        # unzip to proper path
        with zipfile.ZipFile(myfile.name, 'r') as zip_ref:
            zip_ref.extractall(f"{CURRENT_DIR}/../src")

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

        os.remove("reports.zip")
        os.remove("files.zip")
        shutil.rmtree(f"{CURRENT_DIR}/../src/raporlar")
        shutil.rmtree(f"{CURRENT_DIR}/../src/files")
        return response

    return render(request, 'test_calculator/index.html')
