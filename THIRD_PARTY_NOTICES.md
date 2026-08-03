# Third-Party Notices

ClimateGuard Mobile is released under the **MIT License** (see `LICENSE`).
The MIT license covers the project's own source code. Bundled third-party
libraries, fonts, and external data sources retain their own licenses, listed
below.

## Bundled libraries (embedded in `index.html`)

| Component | Purpose | License |
|-----------|---------|---------|
| [Leaflet](https://leafletjs.com/) 1.9.4 | Interactive map | BSD-2-Clause |
| [Leaflet.heat / simpleheat](https://github.com/Leaflet/Leaflet.heat) | Heatmap layer | BSD-2-Clause (© Vladimir Agafonkin) |
| [Remix Icon](https://remixicon.com/) | UI icons (subset) | Apache-2.0 |
| Poppins, Open Sans, IBM Plex Mono | Fonts (subset) | SIL Open Font License 1.1 |

Map tiles are loaded at runtime from OpenStreetMap (© OpenStreetMap
contributors, ODbL) and Esri World Imagery (© Esri, Maxar and others),
subject to each provider's terms.

## Biodiversity data sources

- **iNaturalist** — Reference photos and observation data are fetched at
  runtime from the iNaturalist API. Photos are licensed **individually by each
  author** (default CC BY-NC; also CC0, CC BY, or All Rights Reserved).
  Attribution shown with each photo must be preserved. Observation data on the
  iNaturalist GBIF dataset defaults to CC BY-NC 4.0.
  See <https://www.inaturalist.org/pages/terms>.
- **GBIF** — Taxonomic class information used to build the bundled species
  list is derived from the GBIF Backbone Taxonomy (**CC BY 4.0**). GBIF-mediated
  occurrence data is licensed CC0 / CC BY / CC BY-NC per publisher.
  See <https://www.gbif.org/terms>.
- **IUCN Red List** — The bundled species list and threat categories for the
  Dominican Republic are derived from the IUCN Red List of Threatened Species,
  used under the IUCN Red List Terms of Use (citation required).
  See <https://www.iucnredlist.org/terms/terms-of-use>.

## Institutional partners

Developed for a biodiversity survey initiative supported by KOICA, KNPS
(Korea National Park Service), and the Ministry of Environment and Natural
Resources of the Dominican Republic (MMARN). This is an unofficial
demonstration prototype and is not an official government product.
