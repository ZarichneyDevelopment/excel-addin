#!/usr/bin/env python3
import argparse
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET


NS_MAIN = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
NS_REL = {"rel": "http://schemas.openxmlformats.org/package/2006/relationships"}


def safe_dir_name(name):
    return re.sub(r"[^A-Za-z0-9._-]+", "_", name).strip("_") or "sheet"


def tsv_escape(value):
    if value is None:
        return ""
    return str(value).replace("\\", "\\\\").replace("\t", "\\t").replace("\r", "\\r").replace("\n", "\\n")


def parse_shared_strings(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    xml = ET.fromstring(z.read("xl/sharedStrings.xml"))
    strings = []
    for si in xml.findall("main:si", NS_MAIN):
        # Shared strings can be plain <t> or rich text with <r><t>
        texts = []
        t = si.find("main:t", NS_MAIN)
        if t is not None and t.text:
            texts.append(t.text)
        for r in si.findall("main:r", NS_MAIN):
            rt = r.find("main:t", NS_MAIN)
            if rt is not None and rt.text:
                texts.append(rt.text)
        strings.append("".join(texts))
    return strings


def col_to_index(cell_ref):
    col = 0
    for ch in cell_ref:
        if not ch.isalpha():
            break
        col = col * 26 + (ord(ch.upper()) - ord("A") + 1)
    return col


def parse_workbook(z):
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    sheets = []
    for sheet in wb.findall("main:sheets/main:sheet", NS_MAIN):
        sheets.append(
            {
                "name": sheet.attrib["name"],
                "sheet_id": sheet.attrib.get("sheetId"),
                "rel_id": sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"),
            }
        )

    defined_names = []
    defined = wb.find("main:definedNames", NS_MAIN)
    if defined is not None:
        for dn in defined.findall("main:definedName", NS_MAIN):
            defined_names.append(
                {
                    "name": dn.attrib.get("name"),
                    "text": (dn.text or "").strip(),
                    "local_sheet_id": dn.attrib.get("localSheetId"),
                }
            )

    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rel_map = {
        rel.attrib["Id"]: "xl/" + rel.attrib["Target"]
        for rel in rels.findall("rel:Relationship", NS_REL)
        if rel.attrib.get("Type", "").endswith("/worksheet")
    }

    for sheet in sheets:
        sheet["path"] = rel_map.get(sheet["rel_id"])

    return sheets, defined_names


def parse_tables(z, sheets):
    tables = []
    for sheet in sheets:
        ws_path = sheet.get("path")
        if not ws_path:
            continue
        rels_path = ws_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
        if rels_path not in z.namelist():
            continue
        ws_rels = ET.fromstring(z.read(rels_path))
        for rel in ws_rels.findall("rel:Relationship", NS_REL):
            if not rel.attrib.get("Type", "").endswith("/table"):
                continue
            table_path = "xl/" + rel.attrib["Target"].lstrip("../")
            table_xml = ET.fromstring(z.read(table_path))
            tables.append(
                {
                    "sheet": sheet["name"],
                    "name": table_xml.attrib.get("name"),
                    "display_name": table_xml.attrib.get("displayName"),
                    "ref": table_xml.attrib.get("ref"),
                }
            )
    return tables


def parse_sheet(z, sheet_path, shared_strings):
    xml = ET.fromstring(z.read(sheet_path))
    dimension = xml.find("main:dimension", NS_MAIN)
    dimension_ref = dimension.attrib.get("ref") if dimension is not None else None

    merges = []
    merge_cells = xml.find("main:mergeCells", NS_MAIN)
    if merge_cells is not None:
        for merge in merge_cells.findall("main:mergeCell", NS_MAIN):
            merges.append(merge.attrib.get("ref"))

    cells = []
    sheet_data = xml.find("main:sheetData", NS_MAIN)
    if sheet_data is not None:
        for row in sheet_data.findall("main:row", NS_MAIN):
            for cell in row.findall("main:c", NS_MAIN):
                cell_ref = cell.attrib.get("r")
                cell_type = cell.attrib.get("t")
                style = cell.attrib.get("s")
                value = None
                formula = None

                f = cell.find("main:f", NS_MAIN)
                if f is not None and f.text is not None:
                    formula = "=" + f.text

                v = cell.find("main:v", NS_MAIN)
                if v is not None and v.text is not None:
                    if cell_type == "s":
                        idx = int(v.text)
                        value = shared_strings[idx] if idx < len(shared_strings) else v.text
                    else:
                        value = v.text

                if cell_type == "inlineStr":
                    is_node = cell.find("main:is", NS_MAIN)
                    if is_node is not None:
                        t = is_node.find("main:t", NS_MAIN)
                        if t is not None and t.text is not None:
                            value = t.text

                if cell_ref:
                    cells.append(
                        {
                            "cell": cell_ref,
                            "row": int(re.sub(r"[^0-9]", "", cell_ref) or "0"),
                            "col": col_to_index(cell_ref),
                            "type": cell_type,
                            "style": style,
                            "value": value,
                            "formula": formula,
                        }
                    )

    return {
        "dimension": dimension_ref,
        "merges": merges,
        "cells": cells,
    }


def write_tsv(path, rows, headers):
    with open(path, "w", encoding="utf-8") as f:
        f.write("\t".join(headers) + "\n")
        for row in rows:
            f.write("\t".join(tsv_escape(row.get(h)) for h in headers) + "\n")


def main():
    parser = argparse.ArgumentParser(description="Unpack XLSX to LLM-friendly JSON/TSV outputs.")
    parser.add_argument("xlsx_path", help="Path to .xlsx file")
    parser.add_argument("output_dir", help="Directory to write unpacked outputs")
    args = parser.parse_args()

    os.makedirs(args.output_dir, exist_ok=True)
    sheets_dir = os.path.join(args.output_dir, "sheets")
    os.makedirs(sheets_dir, exist_ok=True)

    with zipfile.ZipFile(args.xlsx_path) as z:
        shared_strings = parse_shared_strings(z)
        sheets, defined_names = parse_workbook(z)
        tables = parse_tables(z, sheets)

        sheet_summaries = []
        for sheet in sheets:
            if not sheet.get("path"):
                continue
            safe_name = safe_dir_name(sheet["name"])
            dir_name = f'{sheet["sheet_id"]}_{safe_name}'
            sheet_dir = os.path.join(sheets_dir, dir_name)
            os.makedirs(sheet_dir, exist_ok=True)

            parsed = parse_sheet(z, sheet["path"], shared_strings)

            with open(os.path.join(sheet_dir, "meta.json"), "w", encoding="utf-8") as f:
                json.dump(
                    {
                        "name": sheet["name"],
                        "sheet_id": sheet["sheet_id"],
                        "dimension": parsed["dimension"],
                        "merge_count": len(parsed["merges"]),
                        "cell_count": len(parsed["cells"]),
                    },
                    f,
                    indent=2,
                )

            write_tsv(
                os.path.join(sheet_dir, "cells.tsv"),
                parsed["cells"],
                ["cell", "row", "col", "type", "style", "value", "formula"],
            )
            write_tsv(
                os.path.join(sheet_dir, "merges.tsv"),
                [{"ref": ref} for ref in parsed["merges"]],
                ["ref"],
            )

            sheet_summaries.append(
                {
                    "name": sheet["name"],
                    "sheet_id": sheet["sheet_id"],
                    "path": sheet["path"],
                    "dir_name": dir_name,
                    "dimension": parsed["dimension"],
                    "cell_count": len(parsed["cells"]),
                    "merge_count": len(parsed["merges"]),
                }
            )

    with open(os.path.join(args.output_dir, "summary.json"), "w", encoding="utf-8") as f:
        json.dump(
            {
                "xlsx_path": args.xlsx_path,
                "sheets": sheet_summaries,
                "defined_names": defined_names,
                "tables": tables,
            },
            f,
            indent=2,
        )

    with open(os.path.join(args.output_dir, "named_ranges.json"), "w", encoding="utf-8") as f:
        json.dump(defined_names, f, indent=2)

    with open(os.path.join(args.output_dir, "tables.json"), "w", encoding="utf-8") as f:
        json.dump(tables, f, indent=2)


if __name__ == "__main__":
    main()
