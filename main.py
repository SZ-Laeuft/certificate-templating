import pathlib
import pprint
import time

import pythoncom
from docxtpl import DocxTemplate
import docx2pdf
from psycopg.rows import class_row
from pydantic import BaseModel, FilePath, DirectoryPath
from pydantic_settings import BaseSettings, SettingsConfigDict
import psycopg
from concurrent.futures import ThreadPoolExecutor
from PyPDF2 import PdfMerger


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file='.env', env_file_encoding='utf-8')
    db_host: str = "localhost"
    db_user: str = "admin"
    db_password: str = "admin"
    db_name: str = "postgres"
    db_port: int = 5432
    output_dir: DirectoryPath = "output/"
    template_path: FilePath = "urkunde-mit-kranz.docx"
    winner_template_path: FilePath = "urkunde-mit-kranz-green.docx"


class User(BaseModel):
    uid: int
    firstname: str
    lastname: str
    count: int
    place: int = 0
    school_class: str = ""


def generate_document(u: User) -> None:
    print(f"User: {u.firstname} {u.lastname} with UID: {u.uid} has {u.count} rounds")
    doc = DocxTemplate("urkunde-mit-kranz-green.docx")
    context = {'name': u.firstname + " " + u.lastname,
               'place': u.place,
               'count': u.count}
    doc.render(context)
    pathlib.Path(f"{settings.output_dir}/{u.school_class}").mkdir(parents=True, exist_ok=True)
    doc.save(f"{settings.output_dir}/{u.school_class}/{u.uid}.docx")
    pythoncom.CoInitialize()
    docx2pdf.convert(f"{settings.output_dir}/{u.school_class}/{u.uid}.docx", f"{settings.output_dir}/{u.school_class}/{u.uid}.pdf")


settings = Settings()
db_url = f"postgresql://{settings.db_user}:{settings.db_password}@{settings.db_host}:{settings.db_port}/{settings.db_name}"
with psycopg.connect(db_url) as conn:
    with conn.cursor(row_factory=class_row(User)) as cur:
        stmt = """
               SELECT DISTINCT u.uid, u.firstname, u.lastname, COUNT(r.scantime), u.school_class
               FROM userinformation u
                        JOIN rounds r ON u.uid = r.uid
               GROUP BY u.uid, u.firstname, u.lastname, u.school_class
               ORDER BY COUNT(r.scantime) DESC;
               """
        cur.execute(stmt)
        users = cur.fetchall()

print("Received following Users from DB:")
pprint.pprint(users)

print("Calculating placings...")
place = 1
current_round_count = None
for user in users:
    if current_round_count is None or user.count == current_round_count:
        user.place = place
    else:
        place += 1
        user.place = place
    current_round_count = user.count

users.sort(key=lambda user: user.school_class)
print("Users with placing and sorted by school class:")
pprint.pprint(users)

start_time = time.time()
print("Generating documents...")
with ThreadPoolExecutor() as executor:
    for user in users:
        executor.submit(generate_document, user)

print("Finished after", time.time() - start_time, "seconds")

# combining each class to one file
for school_class in set(user.school_class for user in users):
    print(f"Processing class: {school_class}")

    # combining all pdfs
    pdfs = list(pathlib.Path(settings.output_dir).glob(f"{school_class}/*.pdf"))
    if not pdfs:
        print(f"No PDFs found for class {school_class}, skipping...")
        continue

    print(f"Combining PDFs for class {school_class}: {[pdf.name for pdf in pdfs]}")
    merger = PdfMerger()

    for pdf in pdfs:
        merger.append(pdf)

    merger.write("result.pdf")
    merger.close()

    doc = DocxTemplate("trennseite.docx")
    context = {'school_class': school_class}
    doc.render(context)
    pathlib.Path(f"{settings.output_dir}/{school_class}").mkdir(parents=True, exist_ok=True)
    doc.save(f"{settings.output_dir}/{school_class}/trennseite.docx")
    pythoncom.CoInitialize()
    docx2pdf.convert(f"{settings.output_dir}/{school_class}/trennseite.docx", f"{settings.output_dir}/{school_class}/trennseite.pdf")

    merger = PdfMerger()
    merger.append(f"{settings.output_dir}/{school_class}/trennseite.pdf")
    merger.append("result.pdf")
    merger.write(f"{settings.output_dir}/{school_class}/combined.pdf")
    merger.close()

pdfs = list(pathlib.Path(settings.output_dir).glob(f"*/combined.pdf"))
merger = PdfMerger()
for pdf in pdfs:
    merger.append(pdf)
merger.write(f"{settings.output_dir}/all_classes_combined.pdf")


