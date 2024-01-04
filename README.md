# pyFaktura

Skrypt tworzący fakturę VAT w formacie PDF.

## Przygotowanie środowiska

1.  Utworzenie wirtualnego środowiska python

    > Poniższe polecenia wykonać w katalogu projektu
    
    ```bash
    pipenv --python 3
    ```

1. Instalacja zależności

    ```bash
    pipenv shell
    pipenv install
    ```
   
## Uruchomienie

> Przed uruchomieniem stworzyć plik konfiguracyjny na podstawie pliku `config.yaml.example`

```bash
python3 pyFaktura.py -c config.yaml
```
