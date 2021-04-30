# versions

# Opis ogólny
Program słuzy do generowania tabeli wybranego oprogramowania z wersjami. 
Tabela jest zapisywana do pliku MS Word.

Program ma zastosowanie m. in. w Księgowości przy raportowaniu używanego oprogramowania księgowego.

# Plik konfiguracyjny config.yaml
W pliku konfiguracyjnym definiuje się lokalizacje plików. Ścieżki mogą prowadzić do lokalnego dysku lub udziału sieciowego.

W pliku można zdefiniować format tabeli. Dostępne formaty są zaprezentowane w pliku [style.docx](doc/style.docx) oraz na https://python-docx.readthedocs.io/en/latest/user/styles-understanding.html#built-in-styles.

Plik config.yaml musi być zapisany jako UTF-8.

# Uwagi
Program działa stosunkowo wolno, w szczególności gdy badane pliki znajdują sie na udziale sieciowym.
