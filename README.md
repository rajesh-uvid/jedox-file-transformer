📂 Jedox Deep Transformer

A powerful Python utility designed to audit and modify Jedox Web Spreadsheets (.wss) and Power BI Connection files (.pb).
🌟 Key Features

    Recursive Processing: Automatically extracts and searches .wss files found inside .pb archives.

    In-Memory Operations: No temporary files are written to disk; all extraction, editing, and re-wrapping happen in RAM.

    Dual-Mode Editing:

        Table Mode: Fast search-and-replace for specific words across all files.

        Manual Mode: Direct XML code editing with a built-in "Pretty Print" formatter.

    Pure XML Export: Option to download a "Flattened" ZIP containing only the raw XML/RELS files from all nested layers.

    CI/CD Ready: Includes automated dependency checks and headless server configurations.

🚀 Quick Start

    Clone the Repo:
    Bash

    git clone https://github.com/username/jedox-deep-transformer.git
    cd jedox-deep-transformer

    Run via Batch (Windows): Double-click run_app.bat. It will automatically verify Python and install streamlit, pandas, and openpyxl globally.

    Run via Terminal:
    Bash

    pip install -r requirements.txt
    streamlit run app.py

🛠️ Project Structure
Plaintext

├── app.py              # Main Streamlit Application
├── run_app.bat         # Global environment launcher
├── requirements.txt    # Python dependencies
└── README.md           # Project documentation

CI/CD Tip

If you are hosting this on a platform like GitHub, the repo name should be lowercase and hyphenated (e.g., jedox-deep-transformer) to follow standard naming conventions.