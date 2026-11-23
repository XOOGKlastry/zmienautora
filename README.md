# ZmienAutora

Hurtowa zmiana autora i śladów edycji w plikach Word i Excel (.docx, .docm, .xlsx, .xlsm i szablony).

**Pobierz:** https://xoogklastry.github.io/zmienautora/

## Co zmienia
- autora, „ostatnio zmodyfikowany przez”, daty, numer wersji, czas edycji, firmę, menedżera
- autora i datę komentarzy oraz śledzonych zmian (Word i Excel)
- konta osób w komentarzach, identyfikatory sesji edycji Worda (rsid)
- ścieżkę szablonu (Word) i lokalną ścieżkę zapisu pliku (Excel)
- usuwa właściwości niestandardowe

Nie zmienia: treści dokumentu, osadzonych obiektów (np. wykresów), makr, dat pliku w Windows.

## Użycie
Uruchom `ZmienAutora.exe`, wybierz folder, wpisz autora, kliknij **Start**.
Oryginały są kopiowane do podfolderu `_kopia_oryginalow`.

Windows może pokazać ostrzeżenie SmartScreen, bo program nie jest podpisany płatnym certyfikatem:
kliknij **Więcej informacji → Uruchom mimo to**. Plik .exe jest budowany automatycznie przez
GitHub Actions z kodu w tym repozytorium (`.github/workflows/release.yml`).

## Uruchomienie ze źródeł
```
python ZmienAutora.py
```
Wymaga tylko Pythona 3.8+ (bez dodatkowych bibliotek).
