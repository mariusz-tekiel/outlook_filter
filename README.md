# outlook_filter
Application checks emails in my Outlook and moves the unneeded ones to the special directory 
Mała, szybka aplikacja do porządkowania skrzynki Outlook (Windows/Office): automatycznie przenosi niechciane oferty pracy i inne maile do wybranych folderów na podstawie czarnej/białej listy, słów kluczowych i reguł. Działa zarówno w trybie jednorazowym (CLI), jak i w tle (zaplanowane uruchomienia, np. Harmonogram zadań).

Cel: odsiać śmieci (np. masowe oferty z portali) i zostawić ważne wiadomości w INBOX.

FUNKCJE:

Czarna lista (blacklist) – słowa/zwroty w temacie, nadawcy lub treści → przenieś do wskazanego folderu (np. Oferty/Automaty).

Biała lista (whitelist) – nadawcy/zwroty, których nigdy nie ruszamy.

Reguły oparte o słowa kluczowe – proste dopasowania lub regex (opcjonalnie).

Filtrowanie „job spam” – predefiniowane frazy typu job, career, hiring, praca, oferta, itp. (można wyłączyć).

Praca na folderach – wybór skrzynki, folderu źródłowego (np. Skrzynka odbiorcza) i docelowych.

Tryb „na sucho” (dry-run) – pokaże, co by przeniósł, bez zmian w Outlooku.

Logi do pliku (CSV/tekst) + podsumowanie w konsoli.

Windows-friendly – współpraca z Outlook (COM przez pywin32).

WYMAGANIA:
Windows z zainstalowanym Microsoft Outlook (konto skonfigurowane).
Python 3.10+

Pakiety:
pywin32

pyyaml (jeśli używasz konfiguracji YAML)
regex (opcjonalnie, jeśli chcesz „prawdziwy” regex; inaczej wystarczy re)

Instalacja paczek:
pip install pywin32 pyyaml regex

Uwaga: Pierwsze użycie pywin32 może wymagać rejestracji modułów:

python -c "import win32com.client"

SZYBKI START:
Skonfiguruj reguły w config.yaml (patrz sekcja niżej).

Uruchom w trybie podglądu (nic nie przenosi):

python outlook_filter.py --dry-run


Jeśli wynik wygląda dobrze – wykonaj faktyczne przenoszenie:

python outlook_filter.py --run
