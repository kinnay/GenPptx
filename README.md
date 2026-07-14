## GenPptx
Dit programma maakt een PowerPoint op basis van een map met afbeeldingen en een `Liederen.txt` bestand dat een lijst van liederen en coupletten bevat. Het programma is gemaakt voor de [Protestantse Gemeente Voorschoten](https://www.dorpskerkvoorschoten.nl/).

### Instructies
Zet een map met afbeeldingen klaar. Deze moeten een van de volgende bestandsnamen hebben:
* `nl-lb-projectie-<liednummer>-muziek-<dia>.png`
* `nl-lb-projectie-<liednummer>-muziek-couplet-<couplet>-<dia>.png`

Bijvoorbeeld `nl-lb-projectie-422-muziek-couplet-1-1.png`.

Zet buiten de map met afbeeldingen een bestand `Liederen.txt` klaar. Iedere regel bevat een liednummer, en eventueel een aantal coupletten door komma's gescheiden, bijvoorbeeld `105 4, 5, 16`. Als er geen coupletten gegeven worden worden alle afbeeldingen van het lied in de PowerPoint gezet.

Sleep de map met afbeeldingen naar het programma toe, totdat "Openen met GenPptx" verschijnt. Zodra je de muisknop loslaat wordt het programma uitgevoerd en maakt het een `Liederen.pptx` bestand.

### Ontwikkeling
Het programma is geschreven in Python 3. De packages die door het programma gebruikt worden kunnen worden geïnstalleerd met: `pip install pillow python-pptx`.

Een executable kan vervolgens worden gemaakt met `pyinstaller --onefile main.py`.

Er kan ook op Linux een Windows executable worden gemaakt door het script `build.sh` uit te voeren. Hiervoor is het nodig dat [Docker](https://docs.docker.com/engine/install/) is geïnstalleerd.
