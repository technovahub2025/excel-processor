from __future__ import annotations

import io
import os
import uuid
from pathlib import Path
from typing import Any

import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
from rapidfuzz import fuzz, process

app = FastAPI(title="Universal Excel Processor API", version="1.0.0")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["http://localhost:5173", "http://127.0.0.1:5173"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# In-memory session store: no database required for this local MVP.
SESSIONS: dict[str, dict[str, Any]] = {}
MAX_UPLOAD_BYTES = 50 * 1024 * 1024
ALLOWED = {".xlsx", ".xlsm", ".xls", ".csv"}


def read_table(content: bytes, filename: str, sheet_name: str | int = 0) -> pd.DataFrame:
    ext = Path(filename).suffix.lower()
    try:
        if ext == ".csv":
            # Try common encodings; pandas will infer delimiters only when configured.
            try:
                return pd.read_csv(io.BytesIO(content), sep=None, engine="python", dtype=object, encoding="utf-8-sig")
            except UnicodeDecodeError:
                return pd.read_csv(io.BytesIO(content), sep=None, engine="python", dtype=object, encoding="latin-1")
        if ext in {".xlsx", ".xlsm", ".xls"}:
            return pd.read_excel(io.BytesIO(content), sheet_name=sheet_name, dtype=object)
    except Exception as exc:
        raise HTTPException(status_code=400, detail=f"Could not read {filename}: {exc}") from exc
    raise HTTPException(status_code=400, detail=f"Unsupported file type: {ext}")


def clean_frame(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() if str(c).strip() else f"Unnamed_{i+1}" for i, c in enumerate(df.columns)]
    df = df.dropna(how="all")
    return df.where(pd.notna(df), None)


def safe_records(df: pd.DataFrame, limit: int = 100) -> list[dict[str, Any]]:
    preview = df.head(limit).copy()
    return preview.astype(object).where(pd.notna(preview), None).to_dict(orient="records")


def matching_key(row: pd.Series, columns: list[str]) -> tuple[str, ...] | None:
    values = tuple(str(row[column]).strip().casefold() if pd.notna(row[column]) else "" for column in columns)
    return values if all(values) else None


@app.get("/")
def root():
    return {"message": "Universal Excel Processor API", "docs": "/docs"}


@app.get("/api/health")
def health():
    return {"status": "ok"}


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    filename = file.filename or "uploaded_file"
    ext = Path(filename).suffix.lower()
    if ext not in ALLOWED:
        raise HTTPException(status_code=400, detail="Upload .xlsx, .xlsm, .xls or .csv files.")
    content = await file.read()
    if len(content) > MAX_UPLOAD_BYTES:
        raise HTTPException(status_code=413, detail="File exceeds the 50 MB MVP upload limit.")
    if not content:
        raise HTTPException(status_code=400, detail="The uploaded file is empty.")

    sheets: list[str] = []
    if ext in {".xlsx", ".xlsm", ".xls"}:
        try:
            excel = pd.ExcelFile(io.BytesIO(content))
            sheets = excel.sheet_names
        except Exception as exc:
            raise HTTPException(status_code=400, detail=f"Invalid Excel workbook: {exc}") from exc
        selected_sheet = sheets[0]
    else:
        sheets = ["CSV"]
        selected_sheet = "CSV"

    df = clean_frame(read_table(content, filename, selected_sheet))
    file_id = str(uuid.uuid4())
    SESSIONS[file_id] = {
        "filename": filename,
        "content": content,
        "sheets": sheets,
        "selected_sheet": selected_sheet,
        "columns": [str(c) for c in df.columns],
        "row_count": int(len(df)),
        "preview": safe_records(df),
    }
    return {
        "file_id": file_id,
        "filename": filename,
        "sheets": sheets,
        "selected_sheet": selected_sheet,
        "columns": [str(c) for c in df.columns],
        "row_count": int(len(df)),
        "preview": safe_records(df),
    }


@app.post("/api/sheet")
async def select_sheet(file_id: str = Form(...), sheet_name: str = Form(...)):
    item = SESSIONS.get(file_id)
    if not item:
        raise HTTPException(status_code=404, detail="Upload session not found. Please upload again.")
    if sheet_name not in item["sheets"]:
        raise HTTPException(status_code=400, detail="Unknown sheet.")
    df = clean_frame(read_table(item["content"], item["filename"], sheet_name))
    item.update({
        "selected_sheet": sheet_name,
        "columns": [str(c) for c in df.columns],
        "row_count": int(len(df)),
        "preview": safe_records(df),
    })
    return {
        "file_id": file_id, "filename": item["filename"], "sheets": item["sheets"],
        "selected_sheet": sheet_name, "columns": item["columns"],
        "row_count": item["row_count"], "preview": item["preview"],
    }


@app.post("/api/process")
async def process_files(
    primary_id: str = Form(...),
    secondary_ids: str = Form(""),
    config_json: str = Form(...),
):
    import json

    primary = SESSIONS.get(primary_id)
    if not primary:
        raise HTTPException(status_code=404, detail="Primary file not found. Upload it again.")
    try:
        config = json.loads(config_json)
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid processing configuration JSON.")

    primary_df = clean_frame(read_table(primary["content"], primary["filename"], primary["selected_sheet"]))
    if primary_df.empty:
        raise HTTPException(status_code=400, detail="Primary file contains no data rows.")

    primary_df = primary_df.reset_index(drop=True)
    primary_df.insert(0, "System_Record_ID", [f"SYS-{i+1:06d}" for i in range(len(primary_df))])

    reports = []
    secondary_list = [x for x in secondary_ids.split(",") if x]
    file_configs = config.get("files", [])

    for sec_id in secondary_list:
        secondary = SESSIONS.get(sec_id)
        if not secondary:
            continue
        sec_cfg = next((x for x in file_configs if x.get("file_id") == sec_id), {})
        left_key = sec_cfg.get("primary_key")
        right_key = sec_cfg.get("secondary_key")
        selected_cols = sec_cfg.get("columns", [])
        mode = sec_cfg.get("mode", "exact")
        threshold = int(sec_cfg.get("threshold", 90))

        sec_df = clean_frame(read_table(secondary["content"], secondary["filename"], secondary["selected_sheet"]))
        if not left_key or not right_key or left_key not in primary_df.columns or right_key not in sec_df.columns:
            reports.append({"filename": secondary["filename"], "error": "Choose valid matching columns.", "matched": 0, "unmatched": len(sec_df)})
            continue

        # Avoid silently matching duplicate keys as unique records.
        left_norm = primary_df[left_key].map(lambda x: str(x).strip().casefold() if x is not None else "")
        right_norm = sec_df[right_key].map(lambda x: str(x).strip().casefold() if x is not None else "")
        left_counts = left_norm[left_norm != ""].value_counts()
        right_counts = right_norm[right_norm != ""].value_counts()
        left_unique = {v for v, n in left_counts.items() if n == 1}
        right_unique = {v for v, n in right_counts.items() if n == 1}
        left_lookup = {v: i for i, v in enumerate(left_norm) if v and v in left_unique}

        match_map: dict[int, int] = {}
        ambiguous = 0
        if mode == "exact":
            for j, value in enumerate(right_norm):
                if value and value in right_unique and value in left_lookup:
                    match_map[j] = left_lookup[value]
                elif value and ((left_counts.get(value, 0) > 1) or (right_counts.get(value, 0) > 1)):
                    ambiguous += 1
        else:
            choices = list(left_unique)
            for j, value in enumerate(right_norm):
                if not value or value in right_unique and value in left_lookup:
                    if value and value in left_lookup:
                        match_map[j] = left_lookup[value]
                    continue
                if not choices:
                    continue
                candidate = process.extractOne(value, choices, scorer=fuzz.WRatio)
                if candidate and candidate[1] >= threshold:
                    candidate_value = candidate[0]
                    if candidate_value in left_lookup and right_counts.get(value, 0) == 1:
                        match_map[j] = left_lookup[candidate_value]

        used_cols = [c for c in selected_cols if c in sec_df.columns and c != right_key]
        # Avoid overwriting primary fields; imported columns get a source prefix.
        for col in used_cols:
            output_name = col if col not in primary_df.columns else f"{Path(secondary['filename']).stem}_{col}"
            primary_df[output_name] = None
            for sec_row, primary_row in match_map.items():
                primary_df.at[primary_row, output_name] = sec_df.at[sec_row, col]

        reports.append({
            "filename": secondary["filename"],
            "matched": len(match_map),
            "unmatched": max(0, len(sec_df) - len(match_map)),
            "ambiguous_or_duplicate": ambiguous,
            "transferred_columns": used_cols,
            "matching_mode": mode,
        })

    result_id = str(uuid.uuid4())
    SESSIONS[result_id] = {"result": primary_df, "filename": "merged_output.xlsx"}
    return {
        "result_id": result_id,
        "columns": [str(c) for c in primary_df.columns],
        "row_count": int(len(primary_df)),
        "preview": safe_records(primary_df),
        "reports": reports,
    }


@app.post("/api/match")
async def match_source_to_target(
    source_id: str = Form(...),
    target_id: str = Form(...),
    config_json: str = Form(...),
):
    import json

    source = SESSIONS.get(source_id)
    target = SESSIONS.get(target_id)
    if not source or not target:
        raise HTTPException(status_code=404, detail="Upload both workbooks again.")
    try:
        config = json.loads(config_json)
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid matching configuration JSON.")

    pairs = config.get("matching_pairs", [])
    transfer_columns = config.get("transfer_columns")
    # Keep accepting the former single-column payload for existing clients.
    if transfer_columns is None and config.get("transfer_column"):
        transfer_columns = [config["transfer_column"]]
    if not isinstance(transfer_columns, list):
        raise HTTPException(status_code=400, detail="Select one or more source transfer columns.")
    # De-duplicate defensively even though the UI prevents duplicate selections.
    transfer_columns = list(dict.fromkeys(transfer_columns))
    source_df = clean_frame(read_table(source["content"], source["filename"], source["selected_sheet"])).reset_index(drop=True)
    target_df = clean_frame(read_table(target["content"], target["filename"], target["selected_sheet"])).reset_index(drop=True)
    if not pairs:
        raise HTTPException(status_code=400, detail="Select at least one matching column pair.")
    if not transfer_columns:
        raise HTTPException(status_code=400, detail="Select at least one source transfer column.")
    if any(not isinstance(column, str) or column not in source_df.columns for column in transfer_columns):
        raise HTTPException(status_code=400, detail="One or more selected transfer columns do not exist in the source workbook.")
    source_columns = [pair.get("source") for pair in pairs]
    target_columns = [pair.get("target") for pair in pairs]
    if any(column not in source_df.columns for column in source_columns) or any(column not in target_df.columns for column in target_columns):
        raise HTTPException(status_code=400, detail="One or more matching columns are invalid.")

    source_lookup: dict[tuple[str, ...], list[int]] = {}
    for index, row in source_df.iterrows():
        key = matching_key(row, source_columns)
        if key is not None:
            source_lookup.setdefault(key, []).append(index)

    output_names: dict[str, str] = {}
    occupied_names = set(target_df.columns)
    for transfer_column in transfer_columns:
        output_name = str(transfer_column)
        if output_name in occupied_names:
            base_name = f"{Path(source['filename']).stem}_{transfer_column}"
            output_name = base_name
            suffix = 2
            while output_name in occupied_names:
                output_name = f"{base_name}_{suffix}"
                suffix += 1
        output_names[transfer_column] = output_name
        occupied_names.add(output_name)

    matched_source_rows: list[int | None] = []
    matched = ambiguous = 0
    for _, row in target_df.iterrows():
        key = matching_key(row, target_columns)
        candidates = source_lookup.get(key, []) if key is not None else []
        if len(candidates) == 1:
            matched += 1
            matched_source_rows.append(candidates[0])
        else:
            if len(candidates) > 1:
                ambiguous += 1
            matched_source_rows.append(None)
    for transfer_column, output_name in output_names.items():
        target_df[output_name] = [
            source_df.at[source_row, transfer_column] if source_row is not None else None
            for source_row in matched_source_rows
        ]

    result_id = str(uuid.uuid4())
    SESSIONS[result_id] = {"result": target_df, "filename": "matched_output.xlsx"}
    unmatched = len(target_df) - matched - ambiguous
    return {
        "result_id": result_id,
        "columns": [str(column) for column in target_df.columns],
        "row_count": int(len(target_df)),
        "preview": safe_records(target_df),
        "matched": matched,
        "unmatched": unmatched,
        "ambiguous": ambiguous,
        "transfer_columns": list(output_names.values()),
        "target_filename": target["filename"],
    }


@app.post("/api/export")
async def export_result(
    result_id: str = Form(...),
    format: str = Form("xlsx"),
    filters_json: str = Form("{}"),
    columns_json: str = Form(""),
):
    import json

    item = SESSIONS.get(result_id)
    if not item or "result" not in item:
        raise HTTPException(status_code=404, detail="Processed result not found. Please process the files again.")
    df = item["result"].copy()
    try:
        filters = json.loads(filters_json or "{}")
    except json.JSONDecodeError:
        raise HTTPException(status_code=400, detail="Invalid filters.")

    for column, value in filters.items():
        if column not in df.columns:
            continue
        column_values = df.loc[:, column]
        # Duplicate headers are filtered consistently across each matching column.
        if isinstance(column_values, pd.DataFrame):
            matches = pd.concat([
                column_values.iloc[:, index].astype(str).str.contains(str(value), case=False, na=False)
                for index in range(column_values.shape[1])
            ], axis=1).all(axis=1) if not isinstance(value, list) else pd.concat([
                column_values.iloc[:, index].astype(str).isin([str(v) for v in value])
                for index in range(column_values.shape[1])
            ], axis=1).all(axis=1)
            df = df[matches]
            continue
        if isinstance(value, list):
            df = df[df[column].astype(str).isin([str(v) for v in value])]
        elif str(value).strip():
            df = df[df[column].astype(str).str.contains(str(value), case=False, na=False)]

    if columns_json:
        try:
            selected_columns = json.loads(columns_json)
        except json.JSONDecodeError:
            raise HTTPException(status_code=400, detail="Invalid export column selection.")
        if not isinstance(selected_columns, list) or not selected_columns:
            raise HTTPException(status_code=400, detail="Select at least one export column.")
        if any(not isinstance(index, int) or isinstance(index, bool) or index < 0 or index >= len(item["result"].columns) for index in selected_columns):
            raise HTTPException(status_code=400, detail="One or more selected export columns are invalid.")
        # De-duplicate positions and retain source dataset order, including repeated header names.
        selected_positions = sorted(set(selected_columns))
        df = df.iloc[:, selected_positions]

    output = io.BytesIO()
    if format.lower() == "csv":
        df.to_csv(output, index=False, encoding="utf-8-sig")
        media_type = "text/csv"
        filename = "processed_data.csv"
    else:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Processed Data")
            ws = writer.sheets["Processed Data"]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for column_cells in ws.columns:
                max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in column_cells)
                ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 40)
        media_type = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        filename = "processed_data.xlsx"
    output.seek(0)
    return StreamingResponse(
        output,
        media_type=media_type,
        headers={"Content-Disposition": f'attachment; filename="{filename}"'},
    )


@app.get("/api/session/{file_id}")
def get_session(file_id: str):
    item = SESSIONS.get(file_id)
    if not item:
        raise HTTPException(status_code=404, detail="Session not found.")
    return {k: item[k] for k in ("filename", "sheets", "selected_sheet", "columns", "row_count", "preview") if k in item}

