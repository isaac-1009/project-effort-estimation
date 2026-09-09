#!/usr/bin/env python3
"""Universal development plan HTML generator.

Reads a JSON config and produces a visually polished HTML page with:
- Gantt chart (week-based grid with month headers)
- Vertical dashed separators every 4 weeks (CSS grid-based)
- Legend with all phases
- Export to PNG button (top-right corner)
- Transparent background support

Usage:
    python scripts/generate_dev_plan.py <config.json> [output.html]

If output path is omitted, writes to data/current/dev-plan.html.

Config JSON structure:
{
    "title": "项目名称",
    "total_weeks": 11,
    "background": "transparent",  // optional: "transparent" or hex color
    "month_labels": ["第1月", "第2月", "第3月"],  // optional
    "legend": [  // optional, auto-generated if omitted
        {"label": "需求阶段", "color": "primary"},
        {"label": "UI 设计", "color": "violet"}
    ],
    "rows": [
        {"label": "需求沟通与评审", "legend": "需求阶段", "start_week": 1, "end_week": 3},
        {"label": "UI设计与评审", "legend": "UI 设计", "start_week": 2, "end_week": 5},
        {"label": "系统发布", "legend": "系统发布", "start_week": 11, "end_week": 11, "is_release": true}
    ]
}

Legend color: primary, secondary, accent, test, violet, rose, sky
Row color: resolved via "legend" field (matches legend label) or "color" field (direct key)
"""

import json
import sys
import math
from pathlib import Path

CSS = """\
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0 }

:root {
  --primary:        #6366F1;
  --primary-dark:   #4F46E5;
  --secondary:      #10B981;
  --secondary-dark: #059669;
  --accent:         #F59E0B;
  --violet:         #8B5CF6;
  --violet-dark:    #7C3AED;
  --rose:           #F43F5E;
  --rose-dark:      #E11D48;
  --sky:            #0EA5E9;
  --sky-dark:       #0284C7;
  --background:     __BG_COLOR__;
  --card-bg:        #FFFFFF;
  --text-primary:   #1E293B;
  --text-secondary: #64748B;
  --text-muted:     #94A3B8;
  --border:         #E2E8F0;
  --shadow:         0 4px 6px -1px rgba(0,0,0,.1), 0 2px 4px -1px rgba(0,0,0,.06);
  --radius:         12px;
}

body {
  font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
  background: var(--background);
  color: var(--text-primary);
  line-height: 1.6;
  -webkit-font-smoothing: antialiased;
}

.container { max-width: 1100px; margin: 0 auto; padding: 2.5rem 2rem 5rem; }

header { text-align: center; margin-bottom: 3rem; position: relative; }
header h1 {
  font-size: 2rem; font-weight: 700; color: var(--text-primary);
  letter-spacing: -.02em; position: relative; display: inline-block; margin-bottom: 1rem;
}
header .export-btn {
  position: absolute;
  top: 0;
  right: 0;
}
header h1::after {
  content: ''; position: absolute; bottom: -8px; left: 50%;
  transform: translateX(-50%); width: 64px; height: 4px;
  background: linear-gradient(90deg, var(--primary), var(--primary-dark)); border-radius: 2px;
}

.sec-label {
  font-size: .7rem; font-weight: 600; letter-spacing: .1em;
  text-transform: uppercase; color: var(--text-muted); margin-bottom: .85rem;
}

.legend { display: flex; gap: 1.25rem; flex-wrap: wrap; margin-bottom: 1.5rem; align-items: center; }
.legend-item { display: flex; align-items: center; gap: .45rem; font-size: .8rem; color: var(--text-secondary); }
.legend-dot { width: 12px; height: 12px; border-radius: 3px; flex-shrink: 0; }

.gantt-card {
  background: var(--card-bg); border-radius: var(--radius);
  box-shadow: var(--shadow); padding: 1.75rem; margin-bottom: 2.5rem; overflow-x: auto;
}
.gantt { min-width: 680px; position: relative; }

.g-month {
  text-align: center; font-size: .78rem; font-weight: 600;
  padding: .35rem 0; color: var(--text-secondary);
  border-bottom: 1px solid var(--border);
  background: #F8FAFC;
}

.g-wcell {
  text-align: center; font-size: .82rem; font-weight: 600;
  padding: .45rem 0; border-bottom: 1px solid var(--border);
}

/* Vertical dashed separators every 4 weeks */
.g-vsep-cell {
  border-right: 2px dashed var(--text-muted);
  opacity: 0.4;
}

.g-row { display: grid; align-items: center; border-bottom: 1px solid var(--border); position: relative; }
.g-row:last-child { border-bottom: none; }

.g-label {
  font-size: .96rem; font-weight: 600; color: var(--text-secondary);
  padding: .45rem .75rem .45rem 0; line-height: 1.3;
  display: flex; align-items: center; justify-content: flex-start; gap: .5rem;
}
.step-num {
  font-size: .82rem; font-weight: 700; color: var(--text-muted);
  flex-shrink: 0; min-width: 20px; text-align: center;
}

.g-bars { position: relative; height: 40px; display: grid; grid-template-columns: repeat(__TOTAL_WEEKS__, 1fr); }
.g-bars-sep { border-right: 2px dashed var(--text-muted); opacity: 0.4; }

.g-bar {
  position: absolute; top: 50%; transform: translateY(-50%); height: 24px;
  border-radius: 6px; display: flex; align-items: center; padding: 0 10px;
  font-size: .82rem; font-weight: 600; white-space: nowrap; overflow: hidden;
  transition: opacity .15s, box-shadow .15s;
}
.g-bar:hover { opacity: .88; box-shadow: 0 2px 8px rgba(0,0,0,.15); }

@media (max-width: 600px) {
  .container { padding: 1.5rem 1rem 4rem; }
  .gantt-card { padding: 1rem; }
}
"""

CDN_SCRIPT = '<script src="https://html2canvas.hertzen.com/dist/html2canvas.min.js"></script>'

EXPORT_SCRIPT = """\
<script>
function exportToPNG() {
  const element = document.getElementById('export-content');
  
  const originalOverflow = document.body.style.overflow;
  document.body.style.overflow = 'visible';
  
  html2canvas(element, {
    backgroundColor: null,
    scale: 2,
    useCORS: true,
    logging: false,
    allowTaint: true
  }).then(canvas => {
    document.body.style.overflow = originalOverflow;
    const link = document.createElement('a');
    link.download = '甘特图___FILENAME__.png';
    link.href = canvas.toDataURL('image/png');
    link.click();
  }).catch(err => {
    document.body.style.overflow = originalOverflow;
    console.error('导出失败:', err);
    alert('导出失败，请重试');
  });
}
</script>"""

COLOR_MAP = {
    "primary":   {"bar": "var(--primary)",   "text": "#fff"},
    "secondary": {"bar": "var(--secondary)", "text": "#fff"},
    "accent":    {"bar": "var(--accent)",    "text": "#fff"},
    "test":      {"bar": "#CBD5E1",          "text": "var(--text-secondary)"},
    "violet":    {"bar": "var(--violet)",    "text": "#fff"},
    "rose":      {"bar": "var(--rose)",      "text": "#fff"},
    "sky":       {"bar": "var(--sky)",       "text": "#fff"},
}

LABEL_COL = "190px"
CIRCLED_NUMS = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫"


def esc(text: str) -> str:
    return (text.replace("&", "&amp;")
               .replace("<", "&lt;")
               .replace(">", "&gt;")
               .replace('"', "&quot;"))


def build_legend_color_map(config: dict) -> dict[str, str]:
    """Map legend label → color key for row lookup."""
    legend = config.get("legend", [])
    return {item["label"]: item.get("color", "primary") for item in legend}


def resolve_row_color(row: dict, legend_color_map: dict[str, str]) -> str:
    """Resolve color key for a row from legend reference or direct color."""
    if "color" in row:
        return row["color"]
    legend_ref = row.get("legend", "")
    return legend_color_map.get(legend_ref, "primary")


def get_bg_color(config: dict) -> str:
    bg = config.get("background", "")
    if bg == "transparent":
        return "transparent"
    return "#f5f5f5"


def build_legend(config: dict) -> str:
    legend_items = config.get("legend")
    if not legend_items:
        seen = set()
        legend_items = []
        for row in config["rows"]:
            ck = row.get("color", "primary")
            if ck not in seen:
                seen.add(ck)
                cm = COLOR_MAP.get(ck, COLOR_MAP["primary"])
                label = row["label"]
                if row.get("is_release"):
                    label = label.replace("🚀 ", "")
                legend_items.append({"label": label, "color_css": cm["bar"]})

    items_html = ""
    for item in legend_items:
        if "color_css" in item:
            color = item["color_css"]
        else:
            ck = item.get("color", "primary")
            cm = COLOR_MAP.get(ck, COLOR_MAP["primary"])
            color = cm["bar"]
        items_html += f'    <div class="legend-item"><div class="legend-dot" style="background:{color};"></div>{esc(item["label"])}</div>\n'

    return f'  <div class="sec-label">图例</div>\n  <div class="legend">\n{items_html}  </div>'


def build_month_headers(config: dict) -> str:
    total = config["total_weeks"]
    num_months = math.ceil(total / 4)
    month_labels = config.get("month_labels")

    cols_html = "<div></div>\n"
    for m in range(num_months):
        w_start = m * 4 + 1
        w_end = min((m + 1) * 4, total)
        gc_start = w_start + 1
        gc_end = w_end + 2
        label = month_labels[m] if month_labels and m < len(month_labels) else f"第{m + 1}月"
        cols_html += f'        <div class="g-month" style="grid-column:{gc_start}/{gc_end};">{esc(label)}</div>\n'

    return (
        f'      <div style="display:grid;grid-template-columns:{LABEL_COL} repeat({total},1fr);margin-bottom:0;">\n'
        f'{cols_html}      </div>'
    )


def build_week_headers(config: dict) -> str:
    total = config["total_weeks"]
    sep_weeks = get_separator_weeks(total)
    cells = "<div></div>\n"
    for w in range(1, total + 1):
        extra_class = " g-vsep-cell" if w in sep_weeks else ""
        cells += f'        <div class="g-wcell{extra_class}">W{w}</div>\n'
    return (
        f'      <div style="display:grid;grid-template-columns:{LABEL_COL} repeat({total},1fr);margin-bottom:.25rem;">\n'
        f'{cells}      </div>'
    )


def get_separator_weeks(total_weeks: int) -> set[int]:
    return {(i + 1) * 4 for i in range(total_weeks // 4)}


def build_gantt_row(row: dict, index: int, total_weeks: int, sep_weeks: set[int], legend_color_map: dict[str, str] | None = None) -> str:
    label = row["label"]
    color_key = resolve_row_color(row, legend_color_map or {})
    cm = COLOR_MAP.get(color_key, COLOR_MAP["primary"])

    if row.get("step_num"):
        num = row["step_num"]
    elif row.get("is_release"):
        num = "🚀"
    else:
        num = CIRCLED_NUMS[index] if index < len(CIRCLED_NUMS) else str(index + 1)

    sw = row["start_week"]
    ew = row["end_week"]
    gap = 0.3
    left_pct = round((sw - 1) / total_weeks * 100 + gap, 1)
    week_span = max(1, ew - sw)
    width_pct = round(week_span / total_weeks * 100 - gap * 2, 1)

    week_cells = ""
    for w in range(1, total_weeks + 1):
        if w in sep_weeks:
            week_cells += '<div class="g-bars-sep"></div>'
        else:
            week_cells += "<div></div>"

    bar_style = f"background:{cm['bar']};color:{cm['text']};left:{left_pct}%;width:{width_pct}%;"

    if row.get("is_release"):
        label_content = f'🚀 {esc(label)}'
    else:
        label_content = f'<span class="step-num">{num}</span>{esc(label)}'

    return (
        f'      <div class="g-row" style="grid-template-columns:{LABEL_COL} 1fr;">\n'
        f'        <div class="g-label">{label_content}</div>\n'
        f'        <div class="g-bars">\n'
        f'          {week_cells}\n'
        f'          <div class="g-bar" style="{bar_style}"></div>\n'
        f'        </div>\n'
        f'      </div>'
    )


def build_gantt(config: dict) -> str:
    total = config["total_weeks"]
    sep_weeks = get_separator_weeks(total)
    legend_color_map = build_legend_color_map(config)
    rows_html = "\n".join(
        build_gantt_row(row, i, total, sep_weeks, legend_color_map)
        for i, row in enumerate(config["rows"])
    )

    months_html = build_month_headers(config)
    weeks_html = build_week_headers(config)

    return (
        f'  <div class="sec-label">排期甘特图（共 {total} 周）</div>\n'
        f'  <div class="gantt-card">\n'
        f'    <div class="gantt">\n'
        f'\n'
        f'\n'
        f'\n'
        f'{months_html}\n\n'
        f'{weeks_html}\n\n'
        f'{rows_html}\n\n'
        f'    </div>\n'
        f'  </div>'
    )


def build_html(config: dict) -> str:
    title = esc(config["title"])
    total = config["total_weeks"]
    bg_color = get_bg_color(config)
    css = CSS.replace('__BG_COLOR__', bg_color).replace('__TOTAL_WEEKS__', str(total))
    export_script = EXPORT_SCRIPT.replace('__FILENAME__', title)

    header_html = (
        f'  <header>\n'
        f'    <h1>{title}</h1>\n'
        f'    <button class="export-btn" onclick="exportToPNG()">📷 导出PNG图片</button>\n'
        f'  </header>'
    )

    parts = [
        '<!DOCTYPE html>\n<html lang="zh-CN">\n<head>',
        '  <meta charset="UTF-8">',
        '  <meta name="viewport" content="width=device-width, initial-scale=1.0">',
        f'  <title>{title}</title>',
        '  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">',
        f'  <style>\n{css}  </style>',
        f'  {CDN_SCRIPT}',
        '</head>\n<body>\n<div class="container">',
        "",
        header_html,
        "",
        "",
        '  <div id="export-content">',
        build_legend(config),
        "",
        "",
        build_gantt(config),
        "  </div>",
        "",
        "</div>",
        export_script,
        "</body>\n</html>",
    ]
    return "\n".join(parts)


def validate_config(config: dict) -> list[str]:
    errors = []
    if "title" not in config:
        errors.append("Missing required field: title")
    if "total_weeks" not in config:
        errors.append("Missing required field: total_weeks")
    if "rows" not in config or not config["rows"]:
        errors.append("Missing required field: rows (must be non-empty)")
    total = config.get("total_weeks", 0)
    for i, row in enumerate(config.get("rows", [])):
        if "label" not in row:
            errors.append(f"Row {i}: missing 'label'")
        if "start_week" not in row or "end_week" not in row:
            errors.append(f"Row {i}: missing 'start_week' or 'end_week'")
        elif row["start_week"] < 1 or row["end_week"] > total or row["start_week"] > row["end_week"]:
            errors.append(f"Row {i}: invalid week range W{row.get('start_week')}-W{row.get('end_week')} (total: {total})")
    return errors


def main():
    import argparse
    parser = argparse.ArgumentParser(
        description="从 config.json 生成开发计划甘特图 HTML"
    )
    parser.add_argument("config_path", help="config.json 路径")
    parser.add_argument("output_path", nargs="?", default=None,
                        help="输出 HTML 路径（可选，默认 data/current/dev-plan.html）")
    args = parser.parse_args()

    config_path = Path(args.config_path)
    if not config_path.exists():
        print(f"Error: config file not found: {config_path}")
        sys.exit(1)

    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)

    errors = validate_config(config)
    if errors:
        print("Config validation errors:")
        for e in errors:
            print(f"  - {e}")
        sys.exit(1)

    html = build_html(config)

    if args.output_path:
        output_path = Path(args.output_path)
    else:
        output_path = Path(__file__).parent.parent / "data" / "current" / "dev-plan.html"

    # 显式 resolve 固定语义：输出路径由操作者通过 CLI 指定（本机自用工具），
    # 但若指向已存在的目录则直接报错，而不是在目录上 open 失败
    output_path = Path(output_path).resolve()
    if output_path.is_dir():
        print(f"Error: output path is a directory: {output_path}")
        sys.exit(1)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Generated: {output_path}")
    print(f"  Title: {config['title']}")
    print(f"  Weeks: {config['total_weeks']}")
    print(f"  Rows:  {len(config['rows'])}")
    print(f"  Background: {get_bg_color(config)}")


if __name__ == "__main__":
    main()
