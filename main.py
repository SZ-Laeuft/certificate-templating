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


class Settings(BaseSettings):
    model_config = SettingsConfigDict(env_file='.env', env_file_encoding='utf-8')
    db_host: str = "localhost"
    db_user: str = "admin"
    db_password: str = "admin"
    db_name: str = "postgres"
    db_port: int = 5432
    output_dir: DirectoryPath = "output/"
    template_path: FilePath = "urkunde.docx"


class User(BaseModel):
    uid: int
    firstname: str
    lastname: str
    count: int
    place: int = 0


def generate_document(u: User) -> None:
    print(f"User: {u.firstname} {u.lastname} with UID: {u.uid} has {u.count} rounds")
    doc = DocxTemplate("urkunde.docx")
    context = {'name': u.firstname + " " + u.lastname,
               'place': u.place,
               'count': u.count}
    doc.render(context)
    doc.save(f"{settings.output_dir}/{u.uid}.docx")
    pythoncom.CoInitialize()
    docx2pdf.convert(f"{settings.output_dir}/{u.uid}.docx", f"{settings.output_dir}/{u.uid}.pdf")


settings = Settings()
db_url = f"postgresql://{settings.db_user}:{settings.db_password}@{settings.db_host}:{settings.db_port}/{settings.db_name}"
with psycopg.connect(db_url) as conn:
    with conn.cursor(row_factory=class_row(User)) as cur:
        stmt = """
               SELECT DISTINCT u.uid, u.firstname, u.lastname, COUNT(r.scantime)
               FROM userinformation u
                        JOIN rounds r ON u.uid = r.uid
               GROUP BY u.uid, u.firstname, u.lastname
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

start_time = time.time()
print("Generating documents...")
with ThreadPoolExecutor() as executor:
    for user in users:
        executor.submit(generate_document, user)
print("Finished after", time.time() - start_time, "seconds")
