# MEMMS — Medical Equipment Maintenance Management System

**Phase 2: Foundation + Equipment Registry**

A desktop application for tracking medical devices and maintenance inside
hospitals. This phase delivers the foundational code only — a working
window, bilingual UI, dark-navy sidebar, and a live dashboard. Equipment registry CRUD is now included; work orders, spare parts, and other modules remain for later phases.

> Phase 1 is **not** production-ready. Do not deploy it in a clinical
> setting.

---

## 1. Tech Stack

| Layer | Choice |
|---|---|
| Language | Python 3.10+ |
| GUI | PySide6 (Qt for Python) |
| Database | SQLite (file under `data/`) |
| ORM | SQLAlchemy 2.x |
| Future DB | PostgreSQL (drop-in swap) |
| Target OS | Windows (Linux/macOS also work) |

---

## 2. Project Structure

```
MEMMS/
├── main.py                # Entry point
├── database.py            # Engine + session factory + init_db()
├── models.py              # SQLAlchemy 2.x declarative models
├── requirements.txt
├── README.md
├── .gitignore
├── data/
│   └── memms.db           # (auto-created on first run, never committed)
└── app/
    ├── __init__.py
    ├── main_window.py     # QMainWindow, sidebar, page stack
    ├── dashboard.py       # 3 live stat cards
    ├── pages.py           # Placeholder pages for future phases
    ├── translations.py    # AR/EN strings + Translator (RTL/LTR)
    └── styles.py          # QSS stylesheet (colors, spacing)
```

> The `app/` package is a small extension to the layout shown in the
> brief. UI code is kept out of the project root so the files a beginner
> edits most (`main.py`, `database.py`, `models.py`) stay small.

---

## 3. Setup on Windows + VS Code

### 3.1 Create a virtual environment

Open a terminal (PowerShell or the VS Code integrated terminal) **inside
the MEMMS project folder**, then run:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks the activation script, run this once and try again:

```powershell
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

On Command Prompt (cmd.exe) use:

```cmd
python -m venv .venv
.venv\Scripts\activate.bat
```

### 3.2 Install dependencies

```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

### 3.3 Run the app

```powershell
python main.py
```

On first run, MEMMS creates `data/memms.db` and all tables. The window
opens in Arabic (RTL) by default.

---

## 4. What Works

- ✅ MEMMS-branded main window with a dark-navy sidebar.
- ✅ Sidebar navigation with 8 sections.
- ✅ Live dashboard: three stat cards backed by real SQL counts.
- ✅ Language switch (Arabic ↔ English) with full RTL ↔ LTR flip.
- ✅ Re-translation of every visible label on switch.
- ✅ Equipment registry: add, list, search, edit, and delete records.
- ✅ Equipment form validation and duplicate asset-code handling.
- ✅ Placeholder pages for sections that ship in later phases.
- ✅ SQLite database created automatically on first launch.
- ✅ Money fields use `Numeric` + `Decimal` (no float arithmetic).
- ✅ Unique constraints on `hospital_code`, `order_number`, `part_number`.

## 5. What is Deferred to Later Phases

- CRUD screens for Equipment, Work Orders, Spare Parts, Staff, etc.
- Preventive maintenance scheduler logic.
- PDF and Excel reports.
- Authentication and permissions.
- Network multi-user mode (PostgreSQL migration is scaffolded but not
  executed in Phase 1).

---

## 6. Manual Test Checklist

Run these after `python main.py` opens the window.

1. **Dashboard shows zeros.** Cards read `—` if the DB is empty; they
   update as soon as records exist.
2. **Open VS Code's SQLite viewer** (or any external tool) and insert a
   row into `equipment`, `work_orders`, and `staff` — then click the
   **تحديث / Refresh** button in the top bar. The three cards should
   jump to the new counts.
3. **Click the language button** (top right). All labels switch between
   Arabic and English, and the layout flips between RTL and LTR
   (sidebar moves to the right in RTL, to the left in LTR).
4. **Click each sidebar entry.** Dashboard is real; the other seven
   sections show a clear placeholder page — buttons are not dead.
5. **Quit and relaunch.** The DB persists in `data/memms.db`.

---

## 7. Common Errors and Fixes

| Error | Cause | Fix |
|---|---|---|
| `ModuleNotFoundError: No module named 'PySide6'` | Dependencies not installed in the active venv | `pip install -r requirements.txt` |
| `qt.qpa.plugin: Could not find the Qt platform plugin "windows"` | Missing Visual C++ runtime on Windows | Install [Microsoft Visual C++ Redistributable](https://aka.ms/vs/17/release/vc_redist.x64.exe) |
| Arabic text shows as squares | System font lacks Arabic glyphs | Windows ships `Segoe UI` and `Tahoma` which both support Arabic. Check the OS language pack. |
| `sqlite3.OperationalError: database is locked` | Another process holds the DB | Close any external SQLite viewer, then retry |
| Layout looks broken after switching language | Sidebar / content side-effect | Click the language button once more — the direction flip is idempotent |

---

## 8. Project Conventions

- Money is stored as `Numeric(12, 2)` and read into Python `Decimal`.
- Every translatable string goes in `app/translations.py` under a stable
  key, never inline in widgets.
- Database access uses the `with get_session() as session:` pattern.
  Sessions always close, even on exceptions.
- All models use the SQLAlchemy 2.x `Mapped[...]` style.

---

## 9. License

This is a college project. Use it freely for learning; do not deploy it
in a real hospital.
