# search_phrase_in_file
Aplikacja do przeszukiwania zawartości plików (Word, Excel, PowerPoint, PDF) z interfejsem graficznym. Umożliwia wyszukiwanie tekstu z uwzględnieniem wielkości liter w wielu plikach jednocześnie.

# File Content Search

Aplikacja desktopowa do przeszukiwania zawartości plików biurowych z interfejsem graficznym.

## Funkcjonalności

- Przeszukiwanie plików w wybranym katalogu i podkatalogach
- Obsługiwane formaty:
  - Word (.doc, .docx)
  - Excel (.xls, .xlsx) 
  - PowerPoint (.ppt, .pptx)
  - PDF (.pdf)
- Możliwość wyboru typów przeszukiwanych plików
- Opcja wyszukiwania z uwzględnieniem wielkości liter
- Okno postępu z aktualnymi statystykami
- Zapis wyników do pliku tekstowego
- Pomijanie plików tymczasowych
- Obsługa błędów dostępu do plików

## Wymagania

- Python 3.x
- Biblioteki:
  - python-docx
  - pandas
  - python-pptx
  - PyPDF2
  - tkinter (standardowa biblioteka Python)

## Instalacja wymaganych bibliotek

```bash
pip install python-docx pandas python-pptx PyPDF2
