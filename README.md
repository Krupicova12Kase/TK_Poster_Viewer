# Prohlížeč posterů na TK

Tento Python program vezme všechny čtyři části posteru na Tvořivou Klávesnici a spojí je do jednoho .png souboru, který následně zobrazí. Konverze do PNG trbá jen pár sekund a s originálními .pptx soubory se nic nestane, vše se uloží do složky output.

## Návod na použití

### Instalace Pythonu a knihoven

1. Nainstalujte [Python](https://www.python.org/downloads/), program byl vytvořen a testován na verzi 3.14
2. Otevřte terminál (Na Windows stiskněte Win+R, napiště cmd a stiskněte Enter) pro nainstalování potřebných knihoven pomocí pip
3. Do terminálu napiště `pip install --upgrade pypiwin32` pro instalování knihovny win32con
4. Nakonec napiště `pip install --upgrade pillow` pro instalování knihovny Pillow, terminál můžete zavřít

### Použití programu

1. Stáhněte soubor `main.py` a umístěte ho do složky mimo ostatní soubory (Mohlo by se stát, že by program použil jiné soubory místo posterů)
2. Vložte všechny čtyři části posteru (hlavičku, levou část, pravou část a střední část) do stejné složky. Soubory nemusí být pojmenované podle příručky, program vezme vše co má koncovku .pptx
3. Spusťte `main.py`. Program by měl začít převádět prezentace na PNG soubory, které následně uloží do složky output a nakonec spojí do jednoho.
4. Po několika sekundách by se měl ukázat celý poster

#### Nastavení

Na začátku souboru je pod komentářem `#settings` nastavení.

1. `close_powerpoint` - Když je hodnota nastavena na `True`, tak zavře aplikaci powerpoint po každém použití. Pokud aktivně pracujete s aplikací, je lepší toto nastavit na `False`, neboli vypnuto. Pokud program nezavře powerpoint po použití, tak bude stále aktivní, což by mohlo způsobit problémy.
