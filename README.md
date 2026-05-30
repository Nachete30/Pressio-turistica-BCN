# Pressió turística i accés al lloguer als barris de Barcelona

Aquest repositori conté els fitxers utilitzats per desenvolupar una visualització interactiva sobre la pressió turística i l'accés al lloguer als barris de Barcelona.

La visualització publicada es pot consultar aquí:

https://public.flourish.studio/story/3693435/

## Objectiu del projecte

L'objectiu és analitzar la pressió residencial als barris de Barcelona combinant diferents dimensions:

- preu mitjà del lloguer,
- lloguer mitjà per metre quadrat,
- renda disponible per persona,
- població resident,
- habitatges d'ús turístic,
- evolució temporal del lloguer.

A partir d'aquestes dades s'han creat indicadors derivats per comparar els barris i detectar possibles zones de tensió residencial.

## Fonts de dades

Les dades utilitzades provenen de fonts obertes i oficials:

- Generalitat de Catalunya: estadístiques del mercat de lloguer.
- Open Data BCN: habitatges d'ús turístic.
- Open Data BCN: padró municipal de Barcelona.
- Open Data BCN: renda disponible de les llars.
- Ajuntament de Barcelona: unitats administratives i barris.

## Estructura recomanada del repositori

```text
data/raw/
  Fitxers originals descarregats de les fonts oficials.

data/processed/
  Taules generades després del tractament de dades.

code/
  Scripts utilitzats per netejar, transformar i integrar les dades.

docs/
  Documentació, memòria del tractament i guió del vídeo.

visualization/
  Enllaç a la visualització publicada a Flourish.
```

## Tractament de dades

El tractament de dades inclou:

1. lectura dels fitxers originals,
2. neteja i selecció de variables rellevants,
3. unificació de les dades per barri,
4. agregació de dades de població i habitatges d'ús turístic,
5. creació d'indicadors derivats,
6. preparació de taules finals per a Flourish.

Els principals indicadors creats són:

- **HUT per 1.000 habitants**: habitatges d'ús turístic normalitzats segons la població del barri.
- **Ràtio lloguer/renda**: relació aproximada entre el cost anual del lloguer i la renda disponible per persona.
- **Variació interanual del lloguer**: evolució percentual del lloguer respecte al període anterior comparable.
- **Índex de pressió residencial**: indicador sintètic construït a partir de variables normalitzades.

## Visualització

La visualització s'ha creat amb Flourish i segueix una estructura narrativa:

1. portada del projecte,
2. mapa de pressió residencial per barris,
3. rànquing dels barris amb més pressió,
4. comparació entre renda disponible i lloguer,
5. pressió dels habitatges d'ús turístic,
6. evolució del lloguer per districtes,
7. descomposició de l'índex mitjançant radar chart,
8. conclusions finals.

Enllaç directe:

https://public.flourish.studio/story/3693435/

## Limitacions

Els resultats s'han d'interpretar amb prudència perquè les fonts de dades no sempre corresponen exactament al mateix any o període temporal. La renda disponible s'ha utilitzat com a variable estructural de context socioeconòmic. L'índex de pressió residencial és una aproximació analítica pròpia i no representa una mesura oficial.

## Llicència

Aquest projecte es publica sota llicència MIT. Consulteu el fitxer `LICENSE` per a més informació.
