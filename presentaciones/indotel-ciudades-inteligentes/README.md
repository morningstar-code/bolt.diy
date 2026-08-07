# Estrategia Nacional de Ciudades Inteligentes — INDOTEL

Presentación de 12 láminas (16:9) para el Consejo Directivo del INDOTEL sobre el
instrumento estratégico InDiCo Global / OASC / U4SSC.

- `INDOTEL-Estrategia-Nacional-Ciudades-Inteligentes.pptx` — entregable editable
- `INDOTEL-Estrategia-Nacional-Ciudades-Inteligentes.pdf` — versión de lectura

## Contenido

| # | Lámina |
|---|--------|
| 1 | Portada — Estrategia Nacional de Ciudades Inteligentes |
| 2 | La visión + indicadores (>50 municipios, >3M ciudadanos, >100 proyectos) |
| 3 | Quiénes son los actores: InDiCo Global, OASC, U4SSC |
| 4 | InDiCo Global: cooperación digital y estándares |
| 5 | OASC: Ciudades Inteligentes Abiertas y Ágiles |
| 6 | Cómo opera OASC (gobernanza, servicios, proyectos, ITU) |
| 7 | El concepto central: Mecanismos Mínimos de Interoperabilidad |
| 8 | ¿Cómo opera OASC? — versión ejecutiva de 4 pilares |
| 9 | El concepto clave: MIMs |
| 10 | Hoja de ruta INDOTEL en 5 fases |
| 11 | Impacto para la ciudadanía |
| 12 | Decisión estratégica |

Cada lámina lleva notas del orador.

## Regenerar el deck

Los scripts son CommonJS (`.cjs`) porque la raíz del repositorio declara
`"type": "module"`.

```bash
npm install pptxgenjs sharp react react-dom react-icons
cd src
node prep.cjs    # recorta, difumina y comprime las fotos -> src/derived/
node icons.cjs   # rasteriza los 33 iconos de react-icons -> src/icons/
node build.cjs   # arma el .pptx en la carpeta padre
```

`src/assets/` guarda las cinco fotografías de origen. `src/derived/` y
`src/icons/` son generados y no se versionan.

### Notas de implementación

- **Sistema de diseño** en la cabecera de `build.cjs`: navy `0B2C5E` dominante,
  azul `1668B8` para iconografía, `2E9BF0` como acento sobre fondo oscuro.
  Tipografía Calibri (Arial para el logotipo).
- **Velo de la portada** (`veil_left.png`): un rectángulo semitransparente
  dejaba una costura vertical visible; se usa un degradado rasterizado.
- **Foto de la lámina 11**: el asset original trae el chip «AMBIENTE»
  superpuesto sobre el brazo del hombre. `prep.cjs` lo difumina con una máscara
  de bordes suaves, recorta al 53,5 % y funde los cuatro lados; el lienzo se
  extiende con transparencia para que el encuadre de PowerPoint no recorte los
  fundidos.
- **Alfa y JPEG**: sharp conserva el formato de entrada, así que cada
  `toBuffer()` intermedio de esa cadena fuerza `.png()` — sin eso el canal alfa
  se aplana y los fundidos desaparecen.
