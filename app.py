import os
import pandas as pd


OUTPUT_FOLDER = "output"
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def process_asc_file(filepath):
    txt_path = os.path.join(OUTPUT_FOLDER, "ASC_Report.txt")
    excel_path = os.path.join(OUTPUT_FOLDER, "ASC_Report.xlsx")

    with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()

    parsed_rows = []
    other_rows = []
    unparsed_lines = []

    for idx, line in enumerate(lines, start=1):
        raw_line = line.rstrip("\n")
        line = line.strip()

        if not line:
            unparsed_lines.append((idx, raw_line))
            continue

        parts = line.split()

        if len(parts) >= 5 and "," in parts[0]:
            try:
                first = parts[0]
                block_no, symbol_name = first.split(",", 1)

                type1 = parts[1]
                num1 = parts[2]
                type2 = parts[3]
                num2 = parts[4]
                desc = " ".join(parts[5:]) if len(parts) > 5 else ""

                address = f"{type1}{num1}"

                if type2 in ["BOOL", "WORD", "INT"]:
                    comment = f"{num2} {desc}".strip()
                    parsed_rows.append([idx, block_no, symbol_name, address, type2, comment])
                else:
                    other_rows.append([idx, block_no, symbol_name, address, type2, num2, desc])

            except:
                unparsed_lines.append((idx, raw_line))
        else:
            unparsed_lines.append((idx, raw_line))

    # Counts
    count_I = count_Q = count_IW = count_QW = 0
    count_BOOL = count_WORD = 0

    for r in parsed_rows:
        addr = r[3].upper()
        t = r[4].upper()

        if addr.startswith("IW"):
            count_IW += 1
        elif addr.startswith("QW"):
            count_QW += 1
        elif addr.startswith("I"):
            count_I += 1
        elif addr.startswith("Q"):
            count_Q += 1

        if t == "BOOL":
            count_BOOL += 1
        elif t == "WORD":
            count_WORD += 1

    # Write TXT
    with open(txt_path, "w", encoding="utf-8") as f:
        f.write("========== TABLE VIEW (ASC SYMBOL LIST) ==========\n\n")
        f.write(f"Total Lines in File     : {len(lines)}\n")
        f.write(f"Parsed Table Entries    : {len(parsed_rows)}\n")
        f.write(f"Other Tags Entries      : {len(other_rows)}\n")
        f.write(f"Unparsed Lines          : {len(unparsed_lines)}\n\n")

        f.write("========== ADDRESS COUNTS ==========\n")
        f.write(f"Addresses starting with IW   : {count_IW}\n")
        f.write(f"Addresses starting with I    : {count_I}\n")
        f.write(f"Addresses starting with QW   : {count_QW}\n")
        f.write(f"Addresses starting with Q    : {count_Q}\n\n")

        f.write("========== TYPE COUNTS ==========\n")
        f.write(f"Total BOOL Tags : {count_BOOL}\n")
        f.write(f"Total WORD Tags : {count_WORD}\n")
        f.write("\n====================================\n\n")

        if parsed_rows:
            f.write("---- PARSED SYMBOL TABLE DATA ----\n\n")
            f.write(f"{'Line':<6} {'Block':<8} {'SymbolName':<35} {'Address':<12} {'Type2':<6} Comment\n")
            f.write("-" * 140 + "\n")

            for r in parsed_rows:
                f.write(f"{r[0]:<6} {r[1]:<8} {r[2]:<35} {r[3]:<12} {r[4]:<6} {r[5]}\n")

        if other_rows:
            f.write("\n\n========== OTHER TAGS ==========\n\n")
            f.write(f"{'Line':<6} {'Block':<8} {'SymbolName':<35} {'Address':<12} {'Type2':<6} {'No2':<6} Description\n")
            f.write("-" * 140 + "\n")

            for r in other_rows:
                f.write(f"{r[0]:<6} {r[1]:<8} {r[2]:<35} {r[3]:<12} {r[4]:<6} {r[5]:<6} {r[6]}\n")

        if unparsed_lines:
            f.write("\n\n---- UNPARSED LINES ----\n\n")
            f.write("-" * 140 + "\n")
            for ln, txt in unparsed_lines:
                f.write(f"{ln:04d} | {txt}\n")

    # Write Excel
    df_summary = pd.DataFrame({
        "Metric": [
            "Total Lines",
            "Parsed Tags",
            "Other Tags",
            "Unparsed Lines",
            "IW Count",
            "I Count",
            "QW Count",
            "Q Count",
            "BOOL Count",
            "WORD Count"
        ],
        "Value": [
            len(lines),
            len(parsed_rows),
            len(other_rows),
            len(unparsed_lines),
            count_IW,
            count_I,
            count_QW,
            count_Q,
            count_BOOL,
            count_WORD
        ]
    })

    df_parsed = pd.DataFrame(parsed_rows, columns=["Line", "Block", "SymbolName", "Address", "Type2", "Comment"])
    df_other = pd.DataFrame(other_rows, columns=["Line", "Block", "SymbolName", "Address", "Type2", "No2", "Description"])
    df_unparsed = pd.DataFrame(unparsed_lines, columns=["Line", "RawText"])

    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        df_summary.to_excel(writer, sheet_name="Summary", index=False)
        df_parsed.to_excel(writer, sheet_name="Parsed_Tags", index=False)
        df_other.to_excel(writer, sheet_name="Other_Tags", index=False)
        df_unparsed.to_excel(writer, sheet_name="Unparsed_Lines", index=False)

    return txt_path, excel_path


def main():
    print("===============================================")
    print("        ASC SYMBOL TABLE EXTRACTION TOOL        ")
    print("===============================================\n")

    while True:
        filepath = input("Enter ASC file path (or drag & drop file here): ").strip().strip('"')

        if not filepath:
            print("❌ No file path entered.\n")
            continue

        if not os.path.exists(filepath):
            print("❌ File not found. Try again.\n")
            continue

        if not filepath.lower().endswith(".asc"):
            print("❌ Not an .asc file. Try again.\n")
            continue

        print("\nProcessing file... please wait...\n")

        try:
            txt_path, excel_path = process_asc_file(filepath)

            print("✅ Done!")
            print(f"📄 TXT Report Saved   : {txt_path}")
            print(f"📊 Excel Report Saved : {excel_path}")
            print("\nOutput folder: output/\n")

        except Exception as e:
            print("❌ Error occurred while processing:")
            print(str(e))

        choice = input("\nDo you want to process another file? (y/n): ").strip().lower()
        if choice != "y":
            print("\nExiting tool...")
            break


if __name__ == "__main__":
    main()