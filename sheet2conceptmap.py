import pandas as pd
from pathlib import Path

def generate_fsh_from_excel(file_path):
    # Excelファイルを読み込む
    excel_file = pd.ExcelFile(file_path)
    output_dir = Path("dist")
    output_dir.mkdir(exist_ok=True)

    for sheet_name in excel_file.sheet_names:
        # 各シートをデータフレームとして読み込む
        df = pd.read_excel(file_path, sheet_name=sheet_name)

        print(df)

        # ConceptMapのFSHファイルを生成
        fsh_content = []
        fsh_content.append(f"// ConceptMap generated from sheet: {sheet_name}")
        fsh_content.append(f"Instance: {sheet_name}")
        fsh_content.append(f"Usage: #definition")

        fsh_content.append(f"name = {sheet_name}")
        fsh_content.append("status = #active")
        fsh_content.append("sexperimental = false")

        fsh_content.append(f'* group[+].source = "{df.iloc[0, 1]}"')
        fsh_content.append(f'* group[=].target = "{df.iloc[0, 3]}"')

        # HLコードとCDISCコードのマッピングを追加
        for _, row in df.iterrows():
            if len(row) < 2 or pd.isna(row[0]) or pd.isna(row[1]):
                continue
            hl_code = row[0]
            cdisc_code = row[1]
            fsh_content.append(f"* group[=].element[+].code = #{hl_code}")
            fsh_content.append(f"* group[=].element[=].target.code = #{cdisc_code}")
            fsh_content.append(f"* group[=].element[=].target.equivalence = #equivalent")

        # FSHファイルに書き込む
        fsh_file_path = output_dir / f"{sheet_name}.fsh"
        with open(fsh_file_path, "w", encoding="utf-8") as fsh_file:
            fsh_file.write("\n".join(fsh_content))

        print(f"FSHファイルを生成しました: {fsh_file_path}")

if __name__ == "__main__":
    excel_path = "Book.xlsx"  # Excelファイルのパス
    generate_fsh_from_excel(excel_path)