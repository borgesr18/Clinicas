import os
import re
import math
from typing import List, Tuple, Optional

import pandas as pd
import numpy as np


INPUT_FILE = "/workspace/Planilha SOSrim.xlsx"
OUTPUT_FILE = "/workspace/Planilha_SOSrim_Reorganizada.xlsx"


def normalize_header(value: str) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    # Collapse multiple spaces
    text = re.sub(r"\s+", " ", text)
    # Keep accents; just title-case words that are not all-caps acronyms
    # Avoid changing very short tokens (ID, UF, RG, etc.)
    parts: List[str] = []
    for token in text.split(" "):
        if len(token) <= 3 and token.isupper():
            parts.append(token)
        else:
            # Title case but keep words like 'da', 'de', 'do' lower if not first
            titled = token.capitalize()
            parts.append(titled)
    cleaned = " ".join(parts)
    # Remove trailing dots and stray punctuation
    cleaned = cleaned.strip(" .:\t\r\n")
    return cleaned


def is_string_like(x) -> bool:
    if pd.isna(x):
        return False
    return isinstance(x, str)


def is_date_like_series(series: pd.Series) -> bool:
    non_null = series.dropna()
    if non_null.empty:
        return False
    sample = non_null.head(100)
    parsed = pd.to_datetime(sample, errors="coerce", dayfirst=True)
    ratio = parsed.notna().mean() if len(sample) > 0 else 0
    return ratio >= 0.7


def is_numeric_like_series(series: pd.Series) -> bool:
    non_null = series.dropna()
    if non_null.empty:
        return False
    sample = non_null.head(200)
    coerced = pd.to_numeric(sample.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False), errors="coerce")
    ratio = coerced.notna().mean()
    return ratio >= 0.8


def detect_header_row(df_raw: pd.DataFrame, max_lookahead: int = 10) -> int:
    # Try to find the first row that looks like headers: mostly strings and not too many NaNs
    best_row = 0
    best_score = -1.0
    lookahead = min(max_lookahead, len(df_raw))
    for r in range(lookahead):
        row = df_raw.iloc[r]
        non_null_ratio = row.notna().mean()
        string_ratio = row.apply(is_string_like).mean()
        # Score favors stringy rows with enough filled cells
        score = (string_ratio * 0.7) + (non_null_ratio * 0.3)
        # Penalize rows with very few filled cells
        if non_null_ratio < 0.4:
            score -= 0.5
        if score > best_score:
            best_score = score
            best_row = r
    return best_row


def prioritize_columns(columns: List[str]) -> List[str]:
    # Score columns based on common Brazilian finance/admin semantics
    priority_order = [
        ("compet", 95),  # competência
        ("period", 92),
        ("mes", 90),
        ("data", 88),
        ("emiss", 86),
        ("venc", 84),
        ("pag", 82),  # pagamento
        ("doc", 80),
        ("nf", 79),
        ("nota", 78),
        ("contrato", 76),
        ("fornec", 74),
        ("cliente", 73),
        ("razao", 72),
        ("cnpj", 71),
        ("cpf", 70),
        ("descr", 65),
        ("histor", 64),
        ("categoria", 63),
        ("centro", 62),
        ("conta", 60),
        ("banco", 58),
        ("agenc", 56),
        ("num", 55),
        ("valor", 50),
        ("preco", 50),
        ("preço", 50),
        ("total", 48),
        ("saldo", 46),
        ("desconto", 44),
        ("juros", 43),
        ("multa", 42),
        ("status", 40),
        ("situ", 39),
        ("obs", 20),
        ("observ", 20),
        ("nota", 19),
    ]
    score_map = {c: 0 for c in columns}
    for col in columns:
        cl = col.lower()
        score = 0
        for token, token_score in priority_order:
            if token in cl:
                score = max(score, token_score)
        score_map[col] = score
    # Stable sort: higher score first, preserve original order for ties
    sorted_cols = sorted(columns, key=lambda c: (-score_map[c], columns.index(c)))
    return sorted_cols


_unique_untitled_seq = 1


def clean_dataframe(df_raw: pd.DataFrame) -> pd.DataFrame:
    global _unique_untitled_seq
    # Drop fully empty rows/cols first for speed
    df = df_raw.copy()
    df = df.dropna(how="all")
    df = df.dropna(how="all", axis=1)

    if df.empty:
        return pd.DataFrame()

    header_row_index = detect_header_row(df)
    header_row = df.iloc[header_row_index].astype(str)

    # Promote header
    new_columns: List[str] = []
    for idx, col_val in enumerate(header_row.tolist()):
        name = normalize_header(col_val)
        if not name:
            name = f"Coluna {_unique_untitled_seq}"
            _unique_untitled_seq += 1
        new_columns.append(name)

    data = df.iloc[header_row_index + 1 :].copy()
    data.columns = new_columns

    # Remove completely empty columns/rows again post-header
    data = data.dropna(how="all")
    data = data.dropna(how="all", axis=1)

    # Strip whitespace in string cells
    for col in data.columns:
        if data[col].dtype == object:
            data[col] = data[col].astype(str).str.strip()
            # Replace common non-values
            data[col] = data[col].replace({"nan": np.nan, "None": np.nan, "": np.nan})

    # Remove duplicate columns keeping the first occurrence
    _, unique_indices = np.unique([c.lower() for c in data.columns], return_index=True)
    keep_mask = np.zeros(len(data.columns), dtype=bool)
    keep_mask[unique_indices] = True
    data = data.loc[:, keep_mask]

    # Heuristic type coercion: dates and numerics
    for col in list(data.columns):
        series = data[col]
        if is_date_like_series(series):
            parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
            data[col] = parsed
        elif is_numeric_like_series(series):
            normalized = pd.to_numeric(
                series.astype(str).str.replace(".", "", regex=False).str.replace(",", ".", regex=False),
                errors="coerce",
            )
            data[col] = normalized

    # Drop duplicate rows
    if not data.empty:
        data = data.drop_duplicates().reset_index(drop=True)

    # Reorder columns by priority
    data = data.loc[:, prioritize_columns(list(data.columns))]

    # Sort rows by the most relevant date column if present
    if not data.empty:
        date_candidates = [c for c in data.columns if pd.api.types.is_datetime64_any_dtype(data[c])]
        if date_candidates:
            # Rank by name tokens as well
            def date_rank(name: str) -> int:
                n = name.lower()
                if "venc" in n:
                    return 0
                if "comp" in n or "compet" in n:
                    return 1
                if "emiss" in n:
                    return 2
                if "pag" in n:
                    return 3
                if "data" in n:
                    return 4
                return 5
            sort_col = sorted(date_candidates, key=lambda n: (date_rank(n), list(data.columns).index(n)))[0]
            data = data.sort_values(by=sort_col, kind="mergesort", na_position="last").reset_index(drop=True)

    # Keep original column order (after cleaning)
    return data.reset_index(drop=True)


def compute_column_widths(df: pd.DataFrame, max_width: int = 60, min_width: int = 8) -> List[int]:
    widths: List[int] = []
    for col in df.columns:
        header_len = len(str(col))
        if df.empty:
            widths.append(max(min_width, header_len + 2))
            continue
        values_sample = df[col].head(500)
        length_series = values_sample.apply(lambda v: len(str(v)) if not pd.isna(v) else 0)
        max_len = int(length_series.max() if not length_series.empty else 0)
        width = min(max_width, max(min_width, max(header_len + 2, max_len + 1)))
        widths.append(width)
    return widths


def write_dataframe_to_worksheet(writer: pd.ExcelWriter, sheet_name: str, df: pd.DataFrame) -> None:
    df_to_write = df.copy()

    # Write the DataFrame
    df_to_write.to_excel(writer, sheet_name=sheet_name, index=False)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    # Formats
    header_format = workbook.add_format({
        "bold": True,
        "text_wrap": True,
        "valign": "center",
        "fg_color": "#F2F2F2",
        "border": 1,
    })

    date_format = workbook.add_format({"num_format": "dd/mm/yyyy"})
    wrap_format = workbook.add_format({"text_wrap": True})
    currency_format = workbook.add_format({"num_format": "R$ #,##0.00"})
    number_format = workbook.add_format({"num_format": "#,##0.00"})
    percent_format = workbook.add_format({"num_format": "0.00%"})
    center_format = workbook.add_format({"align": "center"})

    # Apply header format
    for col_idx, column_name in enumerate(df_to_write.columns):
        worksheet.write(0, col_idx, column_name, header_format)

    # Auto column widths and specific formats
    col_widths = compute_column_widths(df_to_write)
    for col_idx, (column_name, width) in enumerate(zip(df_to_write.columns, col_widths)):
        col_values = df_to_write[column_name]
        name_lower = column_name.lower()
        # Determine format
        col_format = None
        if pd.api.types.is_datetime64_any_dtype(col_values):
            col_format = date_format
        elif pd.api.types.is_numeric_dtype(col_values):
            if any(k in name_lower for k in ["valor", "preço", "preco", "total", "saldo", "pago", "parcela", "juros", "multa", "taxa"]):
                col_format = currency_format
            elif "%" in name_lower or "percent" in name_lower:
                col_format = percent_format
            else:
                col_format = number_format
        else:
            # Wrap text for potentially long text columns
            if width >= 30 or any(k in name_lower for k in ["descr", "observ", "endereço", "endereco", "nota", "coment"]):
                col_format = wrap_format
        if col_format is not None:
            worksheet.set_column(col_idx, col_idx, width, col_format)
        else:
            worksheet.set_column(col_idx, col_idx, width)

        # Center align typical ID-like columns
        if any(k in name_lower for k in ["id", "cpf", "cnpj", "nf", "doc", "nº", "no.", "n.", "numero", "número"]):
            worksheet.set_column(col_idx, col_idx, width, center_format)

    # Freeze header row
    worksheet.freeze_panes(1, 0)

    # Add autofilter
    last_row = len(df_to_write)
    last_col = len(df_to_write.columns) - 1
    if last_col >= 0:
        worksheet.autofilter(0, 0, last_row, last_col)

    # Add table styling with totals
    if last_row >= 0 and last_col >= 0:
        table_columns = []
        for col_name in df_to_write.columns:
            col_def = {"header": col_name}
            series = df_to_write[col_name]
            if pd.api.types.is_numeric_dtype(series):
                col_def["total_function"] = "sum"
            table_columns.append(col_def)
        table_options = {"columns": table_columns, "style": "Table Style Light 9", "total_row": True}
        try:
            worksheet.add_table(0, 0, last_row, last_col, table_options)
        except Exception:
            pass


def add_summary_sheet(writer: pd.ExcelWriter, sheet_to_df: List[Tuple[str, pd.DataFrame]]) -> None:
    # Build a compact summary with per-sheet quick stats
    summary_rows: List[List[object]] = []
    for sheet_name, df in sheet_to_df:
        if df is None or df.empty:
            summary_rows.append([sheet_name, 0, 0, "-", "-"])
            continue
        num_rows = len(df)
        num_cols = len(df.columns)
        # Identify candidate categorical columns (low cardinality)
        cat_cols = []
        for col in df.columns:
            nunique = df[col].nunique(dropna=True)
            if nunique > 0 and nunique <= 20:
                cat_cols.append((col, nunique))
        cat_cols_sorted = sorted(cat_cols, key=lambda x: x[1])[:5]
        cats_preview = ", ".join([f"{c}({n})" for c, n in cat_cols_sorted]) if cat_cols_sorted else "-"

        # Identify date columns
        date_cols = [c for c in df.columns if pd.api.types.is_datetime64_any_dtype(df[c])]
        date_cols_text = ", ".join(date_cols) if date_cols else "-"

        summary_rows.append([sheet_name, num_rows, num_cols, cats_preview, date_cols_text])

    summary_df = pd.DataFrame(summary_rows, columns=["Aba", "Linhas", "Colunas", "Categorias (<=20)", "Datas"])
    summary_df.to_excel(writer, sheet_name="Resumo", index=False)

    worksheet = writer.sheets["Resumo"]
    workbook = writer.book

    header_format = workbook.add_format({
        "bold": True,
        "text_wrap": True,
        "valign": "center",
        "fg_color": "#D9E1F2",
        "border": 1,
    })
    for col_idx, col_name in enumerate(summary_df.columns):
        worksheet.write(0, col_idx, col_name, header_format)

    widths = compute_column_widths(summary_df)
    for idx, width in enumerate(widths):
        worksheet.set_column(idx, idx, width)

    worksheet.freeze_panes(1, 0)
    last_row = len(summary_df)
    last_col = len(summary_df.columns) - 1
    worksheet.autofilter(0, 0, last_row, last_col)


def process_workbook(input_path: str, output_path: str) -> Tuple[List[str], List[str]]:
    # Returns (processed_sheets, skipped_sheets)
    xl = pd.ExcelFile(input_path, engine="openpyxl")
    processed_pairs: List[Tuple[str, pd.DataFrame]] = []
    processed_names: List[str] = []
    skipped: List[str] = []

    for sheet in xl.sheet_names:
        try:
            df_raw = xl.parse(sheet_name=sheet, header=None, dtype=object)
            cleaned = clean_dataframe(df_raw)
            if cleaned is None or cleaned.empty or cleaned.columns.size == 0:
                skipped.append(sheet)
            else:
                processed_pairs.append((sheet, cleaned))
                processed_names.append(sheet)
        except Exception as e:
            skipped.append(sheet)

    # Write output
    with pd.ExcelWriter(output_path, engine="xlsxwriter", datetime_format="dd/mm/yyyy", date_format="dd/mm/yyyy") as writer:
        for sheet_name, df in processed_pairs:
            safe_name = sheet_name
            if len(safe_name) > 31:
                safe_name = safe_name[:28] + "..."
            write_dataframe_to_worksheet(writer, safe_name, df)
        add_summary_sheet(writer, processed_pairs)

    return processed_names, skipped


def main() -> None:
    input_path = INPUT_FILE
    output_path = OUTPUT_FILE

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {input_path}")

    processed, skipped = process_workbook(input_path, output_path)
    print("ABAS PROCESSADAS:")
    for s in processed:
        print(f" - {s}")
    if skipped:
        print("ABAS IGNORADAS (vazias ou inválidas):")
        for s in skipped:
            print(f" - {s}")
    print(f"\nArquivo gerado: {output_path}")


if __name__ == "__main__":
    main()