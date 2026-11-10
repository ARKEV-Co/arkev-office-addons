# ARKEV Office add-ons for ONLYOFFICE Docs

ARKEV Office is ARKEV's document editing (Word, Excel, PowerPoint and PDF) inside the ARKEV app. It runs **ONLYOFFICE Docs Community** (version 10.0.0.1), unmodified, as its own service. ONLYOFFICE Docs is © Ascensio System SIA and licensed under the GNU AGPL v3: https://github.com/ONLYOFFICE/DocumentServer

This repository publishes the two parts ARKEV loads into that editor, under the same licence (AGPL-3.0):

- `locale/ckb/`: the Central Kurdish (Sorani, `ckb`) interface files, mounted into the editor's web apps. ONLYOFFICE has no Sorani interface of its own. The strings were machine-written and checked; a few are still English.
- `plugins/arkev-laya/`: the "ARKEV Laya" system plugin. It lets ARKEV's assistant, Laya, read the open document and apply edits live through the ONLYOFFICE plugin API (`callCommand`). It talks only to ARKEV's API (`https://api.arkev.co`) with a per-document token that ARKEV's server puts in the editor's signed config.

ONLYOFFICE's own logo, loader and About stay in the editor, as its licence requires. "ONLYOFFICE" is a trademark of Ascensio System SIA; this project is not affiliated with or endorsed by it.
