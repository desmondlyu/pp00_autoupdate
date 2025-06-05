# pp00_autoupdate

This project provides **MSS Transfer**, a small tool that turns the comment information inside your Excel `MSS` files into twelve columns required by our `PE` system. You simply pick an Excel file, start the process, and the program takes care of reorganising the data for you.

## Requirements

- Python 3.7 or newer.
- The `openpyxl` package (install it with `pip install openpyxl`).

## How to use it

1. Open a command prompt in this folder.
2. Run the program with `python MSS_transfer.py`.
3. Click **Select MSS File** and choose your Excel file.
4. Click **Start Process**. When the dialog shows **Done**, your converted file will be saved next to the original.

## Repository contents

- `MSS_transfer.py` – the main application window.
- `plaintext` – a short note describing a suggested folder layout.
- `README.md` – the document you are reading now.

The goal is to help non-technical colleagues transform MSS annotations without editing any source code. Just follow the steps above and the program will do the work for you.
