# Third-party notices

| Component | Version | Licence | How it is used | Distributed in this package |
| --- | --- | --- | --- | --- |
| anthropics/skills `skills/pptx` (schemas, design of validators/helpers) | commit at integration time | Apache-2.0 (`licenses/Apache-2.0.txt`) | XSD files copied verbatim into `schemas/ooxml/`; Python logic independently re-implemented | schemas: yes (verbatim); code: no (re-implemented) |
| ISO/IEC 29500-4:2016 XSD (via the above) | 2016 | Published by ISO/IEC / Ecma International; redistributed as shipped by anthropics/skills | Structural validation of slides, layouts, masters, notes, presentation.xml | yes |
| PptxGenJS | 4.0.1 (locked) | MIT | Native object generation (`render.js`) | no (`npm ci`) |
| python-pptx | >=1.0.2,<2 | MIT | verify/revise object layer | no (`pip install`) |
| lxml | >=5 | BSD-3-Clause | XML / XSD | no |
| jsonschema | >=4.18,<5 | MIT | plan contract | no |
| Pillow | >=10,<14 | HPND | image probes | no |
| PyMuPDF (optional) | >=1.24 | AGPL-3.0 or Artifex commercial | optional render verification only | no — do not install if AGPL is not acceptable |

Anthropic's own `THIRD_PARTY_NOTICES.md` lists imageio / FFmpeg etc. for other skill scripts (thumbnail/soffice shims). None of that code or those binaries are included here.
