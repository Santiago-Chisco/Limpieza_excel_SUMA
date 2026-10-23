"""Lógica de limpieza de Excel y gestión de archivos dañados."""
from __future__ import annotations

import csv
import io
import os
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

# En Render se apunta a un disco persistente con la variable DAMAGED_DIR
DAMAGED_DIR = Path(os.environ.get("DAMAGED_DIR", "excel_danados"))
LOG_FILE = DAMAGED_DIR / "registro.csv"


# ---------------------------------------------------------------- lectura
class DamagedFileError(Exception):
    """El archivo no se pudo leer o no contiene datos."""


def read_workbook(file_bytes: bytes, filename: str) -> dict[str, pd.DataFrame]:
    """Lee todas las hojas. Lanza DamagedFileError si el archivo está dañado."""
    if not file_bytes:
        raise DamagedFileError("El archivo está completamente vacío (0 bytes).")

    name = filename.lower()
    try:
        if name.endswith(".csv"):
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), dtype=object)
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(file_bytes), dtype=object, encoding="latin-1")
            sheets = {"Hoja1": df}
        else:
            sheets = pd.read_excel(io.BytesIO(file_bytes), sheet_name=None, dtype=object)
    except Exception as exc:  # noqa: BLE001 - cualquier fallo = archivo dañado
        raise DamagedFileError(f"No se pudo leer el archivo: {exc}") from exc

    if not sheets:
        raise DamagedFileError("El libro no contiene hojas.")
    return sheets


# --------------------------------------------------------------- limpieza
def clean_dataframe(
    df: pd.DataFrame,
    *,
    strip_text: bool = True,
    drop_empty_rows: bool = True,
    drop_empty_cols: bool = True,
    drop_duplicates: bool = True,
    keep: str = "first",
) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Devuelve (df_limpio, df_duplicados_eliminados, estadísticas)."""
    stats = {
        "filas_originales": len(df),
        "columnas_originales": df.shape[1],
        "filas_vacias_eliminadas": 0,
        "columnas_vacias_eliminadas": 0,
        "duplicados_eliminados": 0,
    }
    df = df.copy()

    if strip_text:
        for col in df.columns:
            df[col] = df[col].map(lambda v: v.strip() if isinstance(v, str) else v)
    # Celdas con solo espacios / texto vacío -> vacío real
    df = df.replace(r"^\s*$", pd.NA, regex=True)

    if drop_empty_rows:
        before = len(df)
        df = df.dropna(how="all")
        stats["filas_vacias_eliminadas"] = before - len(df)

    if drop_empty_cols:
        before = df.shape[1]
        df = df.dropna(axis=1, how="all")
        stats["columnas_vacias_eliminadas"] = before - df.shape[1]

    removed = df.iloc[0:0]
    if drop_duplicates and len(df):
        mask = df.duplicated(keep=keep)
        removed = df[mask]
        df = df[~mask]
        stats["duplicados_eliminados"] = int(mask.sum())

    df = df.reset_index(drop=True)
    stats["filas_finales"] = len(df)
    stats["columnas_finales"] = df.shape[1]
    return df, removed, stats


def clean_workbook(sheets: dict[str, pd.DataFrame], include_removed: bool = False, **options):
    """Limpia todas las hojas. Devuelve (bytes_xlsx, resumen_df)."""
    cleaned: dict[str, pd.DataFrame] = {}
    removed_all: dict[str, pd.DataFrame] = {}
    summary_rows = []

    for sheet_name, df in sheets.items():
        clean_df, removed, stats = clean_dataframe(df, **options)
        cleaned[sheet_name] = clean_df
        if len(removed):
            removed_all[sheet_name] = removed
        summary_rows.append({"hoja": sheet_name, **stats})

    summary = pd.DataFrame(summary_rows)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        for sheet_name, df in cleaned.items():
            df.to_excel(writer, sheet_name=str(sheet_name)[:31], index=False)
            _autofit(writer.sheets[str(sheet_name)[:31]], df)
        if include_removed:
            for sheet_name, df in removed_all.items():
                name = f"Dup_{sheet_name}"[:31]
                df.to_excel(writer, sheet_name=name, index=False)
                _autofit(writer.sheets[name], df)
    return buffer.getvalue(), summary


def _autofit(ws, df: pd.DataFrame) -> None:
    for i, col in enumerate(df.columns, start=1):
        width = max([len(str(col))] + [len(str(v)) for v in df[col].head(500)]) + 2
        ws.column_dimensions[ws.cell(row=1, column=i).column_letter].width = min(width, 60)


# ------------------------------------------------- detección de errores
EXCEL_ERRORS = {"#N/A", "#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#NUM!", "#NULL!"}
ERROR_COLUMNS = ["hoja", "fila", "columna", "tipo_error", "valor"]
_NUM_RE = re.compile(r"[-+]?\d+([.,]\d+)?([eE][-+]?\d+)?")


def _is_blank(v) -> bool:
    if isinstance(v, str):
        return v.strip() == ""
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def _is_number(v) -> bool:
    if isinstance(v, bool):
        return False
    if isinstance(v, (int, float)) or type(v).__module__ == "numpy":
        return True
    return isinstance(v, str) and bool(_NUM_RE.fullmatch(v.strip()))


def _sheet_errors(sheet_name: str, df: pd.DataFrame, keep: str) -> list[dict]:
    if len(df) == 0 or df.shape[1] == 0:
        return []
    records: list[dict] = []

    def add(pos, col, tipo, valor=""):
        records.append(
            {
                "hoja": sheet_name,
                "fila": None if pos is None else pos + 2,  # +2: encabezado = fila 1
                "columna": str(col),
                "tipo_error": tipo,
                "valor": str(valor)[:80],
            }
        )

    blank = df.apply(lambda c: c.map(_is_blank))
    empty_rows = blank.all(axis=1)
    empty_cols = blank.all(axis=0)

    for pos, is_empty in enumerate(empty_rows):
        if is_empty:
            add(pos, "(toda la fila)", "Fila vacía")
    for col, is_empty in empty_cols.items():
        if is_empty:
            add(None, col, "Columna vacía")

    for col in df.columns:
        if empty_cols[col]:
            continue
        values = [v for v in df[col] if not _is_blank(v)]
        nums = sum(_is_number(v) for v in values)
        words = sum(isinstance(v, str) and not _is_number(v) for v in values)
        numeric_col = nums > 0 and nums / (nums + words) >= 0.6

        for pos, v in enumerate(df[col]):
            if empty_rows.iloc[pos] or not isinstance(v, str):
                continue
            stripped = v.strip()
            if stripped == "":
                if v != "":
                    add(pos, col, "Celda con solo espacios", repr(v))
                continue
            if v != stripped:
                add(pos, col, "Espacios sobrantes", repr(v))
            if stripped.upper() in EXCEL_ERRORS:
                add(pos, col, "Error de Excel", stripped)
            elif numeric_col and not _is_number(stripped):
                add(pos, col, "Texto en columna numérica", stripped)

    # Duplicados (se comparan sin espacios sobrantes y sin filas/columnas vacías)
    norm = df.apply(lambda c: c.map(lambda v: v.strip() if isinstance(v, str) else v))
    norm = norm.mask(blank)
    sub = norm.loc[~empty_rows.values, ~empty_cols.values]
    if sub.shape[1] and len(sub):
        dup = sub.duplicated(keep=keep)
        pos_map = {idx: i for i, idx in enumerate(df.index)}
        for idx in sub.index[dup.values]:
            row = sub.loc[idx]
            shown = " | ".join(str(x) for x in row if not pd.isna(x))
            add(pos_map[idx], "(toda la fila)", "Fila duplicada", shown)

    records.sort(key=lambda r: (r["fila"] is None, r["fila"] or 0))
    return records


def find_errors(sheets: dict[str, pd.DataFrame], keep: str = "first") -> pd.DataFrame:
    """Tabla con todos los errores encontrados en el archivo original (todas las hojas)."""
    records: list[dict] = []
    for sheet_name, df in sheets.items():
        records.extend(_sheet_errors(str(sheet_name), df, keep))
    return pd.DataFrame(records, columns=ERROR_COLUMNS)


# ------------------------------------------------------ archivos dañados
def save_damaged(file_bytes: bytes, filename: str, reason: str) -> Path:
    DAMAGED_DIR.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe = Path(filename).name
    dest = DAMAGED_DIR / f"{stamp}_{safe}"
    dest.write_bytes(file_bytes)

    new_log = not LOG_FILE.exists()
    with LOG_FILE.open("a", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        if new_log:
            writer.writerow(["fecha", "archivo_guardado", "motivo"])
        writer.writerow([datetime.now().isoformat(timespec="seconds"), dest.name, reason])
    return dest


def list_damaged() -> list[dict]:
    if not DAMAGED_DIR.exists():
        return []
    reasons = {}
    if LOG_FILE.exists():
        with LOG_FILE.open(encoding="utf-8") as f:
            for row in csv.DictReader(f):
                reasons[row["archivo_guardado"]] = row["motivo"]
    items = []
    for p in sorted(DAMAGED_DIR.iterdir(), reverse=True):
        if p.name == LOG_FILE.name or not p.is_file():
            continue
        items.append(
            {
                "nombre": p.name,
                "ruta": p,
                "tamano_kb": round(p.stat().st_size / 1024, 1),
                "fecha": datetime.fromtimestamp(p.stat().st_mtime).strftime("%Y-%m-%d %H:%M"),
                "motivo": reasons.get(p.name, ""),
            }
        )
    return items


def delete_damaged(path: Path) -> None:
    if path.exists() and path.parent == DAMAGED_DIR:
        path.unlink()
