# GRUPPINDELNING KLOCKARHAGSSKOLAN

## PROCEDUR
Exportera klass 1A-9C från Infomentor med följande kolumner i följande ordning:

```[Elev Klass, Elev Namn, Elev Grupper, Elev Personnummer]```

Öppna upp i Excel eller liknande och sortera raderna från A->Z.

Klipp och klistra in i elevlista:elevlista

Ta bort eventuella mail som finns under kolumnen ```Elev mail```.

Kör skriptet

```$ python klock_gruppper```

### SKRIPTET I TVÅ STEG

#### FÖRST
Skriptet ser till att rätt mail läggs till rätt elev via referensbladet elevlista:edukonto

När det är klart frågas användaren om programmet skall fortsätta kolla gruppindelningen.

#### ANDRA
Väljer användaren att fortsätta med gruppindelningen generaras csv-filer för de olika grupperna beroende på vilka grupper som finns för eleven i ```Elev Grupper``` kolumnen. Detta med hjälp av ```config.py``` som 'talar om' vilka klasser och grupper som skall kollas upp och skapas.

Det skapas samtidigt excelfiler med samma grupper för att snabbt kunna lägga upp dessa på driven.