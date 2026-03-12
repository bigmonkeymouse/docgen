# DocGen GUI

## Default Template Directory

- Default path: `<script_dir>/word_template`, where `script_dir` is the directory of `template_filler.py` (i.e., `f:\ai_proj\doc_gen\code\docgen\word_template`).
- On first launch, the application ensures this folder exists. If creating this folder fails (e.g., lack of permission), it falls back to the user's home directory: `<USER_HOME>/docgen_templates`.
- All template-related operations read from the configured "Word template folder". The GUI exposes this folder in Preferences and via the “选择模板” button.

## Missing Template Handling

Before generating documents, the app verifies that the template folder exists and contains at least one `.docx` file. If not:

- A modal dialog is shown with a Critical icon titled “模板缺失” and a button to open the parent directory so the user can create/populate the folder.
- The “生成文档” button remains disabled until a valid template is detected.

## UI Compactness

- Window minimum width is 800 px (height 480 px).
- Compact layout with 12 px outer margins and 6 px control spacing.
- The toolbar includes “选择模板”, “生成文档”, and “打开输出目录” on the same row.
- System default font at a slightly smaller size for better use of space.

