from dataclasses import dataclass

@dataclass
class Row:
    row_type: str
    authorships_raw_author_name: str
    mfn_authors: str
    publication_year: str
    title: str
    journal: str
    biblio_volume: str
    biblio_issue: str
    biblio_first_page: str
    biblio_last_page: str
    doi: str
    open_access_is_oa: bool
    editor: str
    book_title: str
    book_serie: str
    print: str
    publisher: str
    dataset: str