"""
analyze_dialogue.py

Reads an Excel file (Sheet 1, columns A=user, B=AI) and identifies up to 3 rows
showing 垂直的な深まり (vertical deepening). Applies blue cell shading and adds
evaluation comments to those cells.

Usage:
    python analyze_dialogue.py <path_to_excel_file.xlsx>
"""

import sys
import json
import anthropic
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

BLUE_FILL = PatternFill(fill_type="solid", fgColor="ADD8E6")


def read_conversation(ws) -> list[dict]:
    rows = []
    for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), start=1):
        a_val = row[0] if len(row) > 0 else None
        b_val = row[1] if len(row) > 1 else None
        if a_val is None and b_val is None:
            continue
        rows.append({"row": row_idx, "user": a_val or "", "ai": b_val or ""})
    return rows


def build_prompt(rows: list[dict]) -> str:
    lines = []
    for r in rows:
        lines.append(f"行{r['row']} ユーザー: {r['user']}")
        lines.append(f"行{r['row']} AI: {r['ai']}")
        lines.append("")

    conversation_text = "\n".join(lines)

    return f"""以下はユーザーとAIの対話記録です（Excelの行番号付き）。

{conversation_text}

この対話の中で「垂直的な深まり」が生まれている箇所を最大3か所特定してください。
垂直的な深まりとは、会話が表面的な情報交換にとどまらず、より深い洞察・理解・気づき・感情的または知的な発展が生まれている瞬間のことです。

以下のJSON形式で回答してください。他のテキストは含めないでください。

{{
  "selections": [
    {{
      "row": <行番号（整数）>,
      "column": "<'A'または'B'のどちらに深まりが顕れているか>",
      "comment": "<その箇所が垂直的な深まりである理由（100〜200字程度の日本語）>"
    }}
  ]
}}

selectionは最大3件です。最も深まりが顕著な箇所を優先してください。"""


def call_claude(conversation_rows: list[dict]) -> list[dict]:
    client = anthropic.Anthropic()
    prompt = build_prompt(conversation_rows)

    response = client.messages.create(
        model="claude-opus-4-8",
        max_tokens=4096,
        thinking={"type": "adaptive"},
        messages=[{"role": "user", "content": prompt}],
    )

    # Extract the text content block
    text_content = ""
    for block in response.content:
        if block.type == "text":
            text_content = block.text
            break

    # Parse JSON from the response
    start = text_content.find("{")
    end = text_content.rfind("}") + 1
    if start == -1 or end == 0:
        raise ValueError(f"JSON not found in Claude response:\n{text_content}")

    data = json.loads(text_content[start:end])
    return data.get("selections", [])


def apply_formatting(ws, selections: list[dict]):
    for sel in selections:
        row = sel.get("row")
        col_letter = sel.get("column", "A").upper()
        comment_text = sel.get("comment", "")

        if not row or col_letter not in ("A", "B"):
            continue

        col_idx = 1 if col_letter == "A" else 2
        cell = ws.cell(row=row, column=col_idx)

        cell.fill = BLUE_FILL

        comment = Comment(comment_text, "Claude")
        comment.width = 300
        comment.height = 120
        cell.comment = comment

        print(f"  行{row} 列{col_letter}: 青色網掛 + コメント付加")


def main():
    if len(sys.argv) < 2:
        print("使い方: python analyze_dialogue.py <Excelファイルのパス>")
        sys.exit(1)

    filepath = sys.argv[1]
    print(f"ファイルを読み込んでいます: {filepath}")

    wb = load_workbook(filepath)
    ws = wb.worksheets[0]

    print("対話データを読み込んでいます...")
    rows = read_conversation(ws)
    if not rows:
        print("対話データが見つかりませんでした。")
        sys.exit(1)
    print(f"  {len(rows)} 行の対話を読み込みました")

    print("Claude に垂直的な深まりを分析させています...")
    selections = call_claude(rows)
    print(f"  {len(selections)} か所の垂直的な深まりを特定しました")

    print("セルを書式設定しています...")
    apply_formatting(ws, selections)

    wb.save(filepath)
    print(f"\n完了: '{filepath}' に保存しました")

    print("\n特定された垂直的な深まり:")
    for i, sel in enumerate(selections, 1):
        print(f"\n[{i}] 行{sel.get('row')} 列{sel.get('column')}")
        print(f"    {sel.get('comment', '')}")


if __name__ == "__main__":
    main()
