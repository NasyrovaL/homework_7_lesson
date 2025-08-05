import os
import zipfile
from io import BytesIO
import pypdf
from openpyxl import load_workbook


def test_create_archive():
    # Создание папки, если не существует
    if not os.path.exists('my_archives'):
        os.mkdir('my_archives')

    archive_path = os.path.join('my_archives', 'archive_file.zip')

    files = [
        'Файл_CSV.csv',
        'Файл_xlsx.xlsx',
        'Файл_PDF.pdf'
    ]

    # Создание ZIP архива и добавление файлов
    with zipfile.ZipFile(archive_path, 'w') as zf:
        for file in files:
            file_path = os.path.join(os.getcwd(), file)
            if os.path.exists(file_path):
                zf.write(file_path, os.path.basename(file_path))

    assert os.path.exists(archive_path), "Архив не был создан"

    with zipfile.ZipFile(archive_path, 'r') as zf:
        archive_contents = zf.namelist()
        for file in files:
            assert file in archive_contents, f"Файл '{file}' не найден в архиве"

#    print("Тест пройден: архив создан, все файлы добавлены.")

 # Чтение содержимого csv файла внутри архива
def test_check_csv():

    archive_path = 'my_archives/archive_file.zip'
    target_file = 'Файл_CSV.csv'

    with zipfile.ZipFile(archive_path, 'r') as zf:
        with zf.open(target_file) as f:
            content = f.read().decode('utf-8')  # если файл в UTF-8
            assert content == 'Тестовый файл csv'


 # Чтение содержимого pdf файла внутри архива
def test_check_pdf():
    archive_path = 'my_archives/archive_file.zip'
    pdf_name = 'Файл_PDF.pdf'
    expected_text = 'Тестовый файл pdf '

    with zipfile.ZipFile(archive_path, 'r') as archive:
        with archive.open(pdf_name) as file:
            pdf_bytes = file.read()
            pdf_stream = BytesIO(pdf_bytes)
            reader = pypdf.PdfReader(pdf_stream)
            for page in reader.pages:
                content = page.extract_text()
                assert content == expected_text


 # Чтение содержимого xlsx файла внутри архива
def test_check_xlsx():
    archive_path = 'my_archives/archive_file.zip'
    xlsx_name = 'Файл_xlsx.xlsx'
    expected_cell_value = 'Тестовый файл xlsx'

    with zipfile.ZipFile(archive_path, 'r') as archive:
        with archive.open(xlsx_name) as file:
            xlsx_bytes = file.read()
            xlsx_stream = BytesIO(xlsx_bytes)

            # Загрузка Excel из потока
            workbook = load_workbook(filename=xlsx_stream, read_only=True)
            sheet = workbook.active
            actual_value = sheet['A1'].value

            assert actual_value == expected_cell_value