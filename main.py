import csv
from docx import Document

cols = {
      0: "quelle",
      1: "FB1",
      2: "FB2",
      3: "FB3",
      4: "GD",
      5: "doi",
      6: "title",
      7: "authorships.raw_author_name",
      8: "mfn-authors",
      9: "publication_year",
     10: "publication_date",
     11: "language",
     12: "type",
     13: "journal",
     14: "publisher",
     15: "primary_location.license",
     16: "open_access.is_oa",
     17: "open_access.oa_status",
     18: "cites_mfn_collection_specimen",
     19: "is_taxonomic_revision",
     20: "is_species_description",
     21: "abteilung 1",
     22: "abteilung 2",
     23: "abteilung 3",
     24: "kommentar",
     25: "biblio.volume",
     26: "biblio.issue",
     27: "biblio.first_page",
     28: "biblio.last_page",
     29: "locations.landing_page_url",
     30: "type_crossref",
     31: "booktitle",
     32: "editor",
     33: "print",
     34: "edition",
     35: "book serie" }

cols_rev = dict((v,k) for k,v in cols.items())

pubtypes =  set()

with open('../Master_cleaned_2024_sp_2025_02_20_UTF-8.CSV', newline='\n') as csvfile:
    spamreader = csv.reader(csvfile, delimiter=';', quotechar='"')
    for row in spamreader:
        for idx, field in enumerate(row):
            if idx == cols_rev.get('type'):
                pubtypes.add(field)

print(pubtypes)

# http://python-docx.readthedocs.io/en/latest/

# document = Document()
# document.add_heading('Document Title', 0)
# document.save('demo.docx')