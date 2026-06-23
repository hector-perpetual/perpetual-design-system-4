# Perpetual Design System 4 — Template "Marketing"

Plantilla de presentación (`.pptx`, 16:9) con el layout estilo "Marketing for your Brand", reinterpretado con la marca Perpetual Technologies: paleta oficial (azul + naranja + amarillo), tipografía **Armin Grotesk**, logos correctos y reglas duras de marca (fondo blanco en claro, nunca cream ni negro puro, logo nunca recoloreado).

## Entregable

- **`dist/perpetual-marketing-template.pptx`** — 13 slides on-brand, editables en PowerPoint o Keynote.

## Las 13 slides

| # | Slide | # | Slide |
|---|---|---|---|
| 1 | Portada | 8 | Herramientas de marketing |
| 2 | Bienvenida | 9 | Opiniones de clientes |
| 3 | Equipo | 10 | Nueva vía del éxito |
| 4 | Trabajo actual | 11 | Plan de estrategia |
| 5 | Lo que ofrecemos | 12 | Crece con nosotros |
| 6 | Tendencias de mercado | 13 | Contáctanos |
| 7 | Nuestras soluciones | | |

## Marca aplicada

- **Color:** azul `#1a56db` + naranja `#f97316` + amarillo `#fbb900`. Cards oscuras en `#0b1220`.
- **Fondo:** blanco (el cream del mockup se reemplaza por blanco, regla de marca).
- **Formas:** "blobs" circulares en azul/naranja/amarillo, como en el mockup.
- **Tipografía:** Armin Grotesk (Black para títulos/números, SemiBold para labels, Regular para cuerpo).
- **Logo:** color en claro, dark sobre fondos azules. Embebido como imagen fiel.
- **Imágenes:** placeholders con glifo de media (reemplazar por fotos reales).
- Contenido placeholder realista en español, sin lorem ipsum ni em-dash.

## Regenerar

```bash
pip install python-pptx
python build.py
```

Reconstruye `dist/perpetual-marketing-template.pptx`. Los logos PNG se rasterizan desde los SVG oficiales del design system.

## Estructura

```
perpetual-design-system-4/
├── build.py                                 # generador (python-pptx)
├── assets/
│   ├── logo/                                # color / dark (SVG + PNG)
│   └── fonts/                               # Armin Grotesk (5 OTF)
└── dist/perpetual-marketing-template.pptx   # entregable
```

> Fuente de verdad de la marca (tokens, reglas): repo [`perpetual-design-system-2`](https://github.com/hector-perpetual/perpetual-design-system-2).

## Marca (fuente de verdad)

La carpeta `brand/` contiene los tokens, componentes y reglas de marca de Perpetual (SKILL.md + references). Este repo es autosuficiente: diseno, fuentes (Armin Grotesk embebida), logos y reglas de marca en un solo lugar.
