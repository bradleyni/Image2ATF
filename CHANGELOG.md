# Changelog

The entries below summarize changes found in successive source versions. The local snapshots were imported into GitHub together.

## Unreleased

- Imported 26 historical source snapshots and documented the changes between versions.

## Version history

The current version is linked from the repository root. Older local snapshots are under `versions/`, with the internal proxy default removed and the embedded contact email removed from V2 through V2.3. These snapshots are for reference and may store API keys in application settings or depend on retired API models.

| Version | Main changes |
| --- | --- |
| [V2.9.17](Image2ATF_V2.9.17.py) | Preserve drawing order for entirely manual masks and use reading order when any automatic mask is present. Keep multiline text during reordering and protect API keys with the system keyring and request headers. |
| [V2.9.16](versions/Image2ATF_V2.9.16.py) | Strip rich-text markup before checking for numeric-only or single-letter snippets. |
| [V2.9.15](versions/Image2ATF_V2.9.15.py) | Normalize CRLF and CR line endings in mask OCR, export ordering, and snippet filtering. |
| [V2.9.14](versions/Image2ATF_V2.9.14.py) | Apply the Word Table Grid style to nested tables in both ATF columns. |
| [V2.9.13](versions/Image2ATF_V2.9.13.py) | Move optional numeric and single-letter filtering to ATF export so the OCR result remains editable. |
| [V2.9.12](versions/Image2ATF_V2.9.12.py) | Recreate detected tables in both ATF columns, hide text in the source copy, and adjust font size for dense tables. |
| [V2.9.11](versions/Image2ATF_V2.9.11.py) | Recognize table markup in mask OCR and render it as editable Word tables. |
| [V2.9.10](versions/Image2ATF_V2.9.10.py) | Extend snippet filtering to single letters and make the option off by default. |
| [V2.9.9](versions/Image2ATF_V2.9.9.py) | Order masked text by its position on the image when exporting an ATF. |
| [V2.9.8](versions/Image2ATF_V2.9.8.py) | Give ATF Creator its own Gemini model choices and use a separate model for orientation detection. |
| [V2.9.7](versions/Image2ATF_V2.9.7.py) | Retry transient OCR failures and prompt before exporting when some files are incomplete. |
| [V2.9.6](versions/Image2ATF_V2.9.6.py) | Keep snippets with decimal separators when filtering numeric-only text. |
| [V2.9.5](versions/Image2ATF_V2.9.5.py) | The local source is byte-for-byte identical to V2.9.4. |
| [V2.9.4](versions/Image2ATF_V2.9.4.py) | Add an option to exclude numeric-only snippets from mask OCR. |
| [V2.9.3](versions/Image2ATF_V2.9.3.py) | Use the font bundled with Pillow to label OCR masks. |
| [V2.9.2](versions/Image2ATF_V2.9.2.py) | Add diagnostic logging around mask OCR and worker failures. |
| [V2.9.1](versions/Image2ATF_V2.9.1.py) | Add crash logs and serialize mask image rendering while keeping network requests parallel. |
| [V2.9](versions/Image2ATF_V2.9.py) | Add a rich-text OCR preview and a duplicate-line option for ATF export. |
| [V2.8](versions/Image2ATF_V2.8.py) | Refine PSD text placement, background generation, and font selection, and derive the ATF filename from the selected OCR engine. |
| [V2.7](versions/Image2ATF_V2.7.py) | Improve rich-text parsing, image preview navigation, and PSD rotation. |
| [V2.6](versions/Image2ATF_V2.6.py) | Improve transparency and background handling in PSD generation and refine Adobe DTP import and export. |
| [V2.5](versions/Image2ATF_V2.5.py) | Add a dedicated PSD mask editor and automatic PSD box detection. |
| [V2.4](versions/Image2ATF_V2.4.py) | Refine image previews, OCR batch controls, and the Adobe workflow. Add access to the user manual from the app. |
| [V2.3](versions/Image2ATF_V2.3.py) | Introduce Image-to-PSD conversion and update the Adobe DTP workflow. |
| [V2.2](versions/Image2ATF_V2.2.py) | Add Adobe DTP import and export with Word and XML support. |
| [V2.1](versions/Image2ATF_V2.1.py) | Refine automatic box grouping and numbered mask OCR parsing, and add image preview zoom. |
| [V2](versions/Image2ATF_V2.py) | Earliest archived source, with image OCR, mask editing, and ATF Word export. |

### Original GitHub V2.8 package

- Initial GitHub repository packaging.
- Includes the PyQt6 script, dependency list, launch scripts, and GitHub upload guide.
- Scrubbed the hard-coded internal proxy default before packaging.

The repository-root [V2.8 script](Image2ATF_V2.8.py) is the original GitHub variant with subsequent keyring and Gemini request-header changes. The table links to the separate local V2.8 snapshot.
