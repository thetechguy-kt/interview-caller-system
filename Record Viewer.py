from flask import Flask, render_template_string
from openpyxl import load_workbook
from datetime import datetime
import os

app = Flask(__name__)

EXCEL_PATH = "candidate_list.xlsx"

def excel_to_html(filepath):
    if not os.path.exists(filepath):
        return "<p style='color:#ffdede'>Excel file not found.</p>"

    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    rows = list(ws.rows)
    if not rows:
        return "<p style='color:#ffdede'>No data found in Excel.</p>"

    # Build table with semantic classes for styling
    html = '<table class="candidate-table" role="table">'
    # Header
    html += '<thead><tr>'
    for cell in rows[0]:
        v = cell.value if cell.value is not None else ""
        html += f'<th scope="col">{v}</th>'
    html += '</tr></thead>'

    # Body
    html += '<tbody>'
    for i, row in enumerate(rows[1:]):
        row_class = "even" if i % 2 == 0 else "odd"
        html += f'<tr class="{row_class}">'
        for j, cell in enumerate(row):
            cell_val = cell.value if cell.value is not None else ""
            # Assume name column is 4th header (index 3). Adjust if needed.
            if j == 3:
                html += f'<td class="name-cell">{cell_val}</td>'
            else:
                html += f'<td>{cell_val}</td>'
        html += '</tr>'
    html += '</tbody></table>'
    return html

@app.route('/')
def index():
    table_html = excel_to_html(EXCEL_PATH)
    now = datetime.now().strftime("%A, %d %B %Y  |  %I:%M:%S %p")
    # Template: dark tech theme, Montserrat (Google Fonts), bold bright time, subtle name highlight
    page = f'''<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>IITC Candidate Record Viewer</title>

<!-- Montserrat from Google Fonts (fallback to system fonts / Aptos if installed) -->
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600;700;800&display=swap" rel="stylesheet">

<style>
  :root{{
    --bg: #0f1214;            /* page background */
    --panel: #15171a;         /* table panel */
    --muted: #bfc8cc;         /* muted text */
    --title: #00e5ff;         /* bright cyan for headings */
    --subtitle: #00bfa5;      /* softer teal for time */
    --accent: #0a84a6;        /* buttons / accents */
    --row-odd: #121416;
    --row-even: #1a1c1e;
    --name-highlight: rgba(0,191,165,0.08); /* subtle name background */
    --border: rgba(255,255,255,0.04);
  }}

  html,body {{
    height:100%;
    margin:0;
    background:var(--bg);
    color:var(--muted);
    font-family: "Montserrat", Aptos, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    -webkit-font-smoothing:antialiased;
    -moz-osx-font-smoothing:grayscale;
  }}

  .container {{
    max-width:1200px;
    margin:32px auto;
    padding:28px;
    background:linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01));
    border-radius:12px;
    box-shadow: 0 8px 30px rgba(0,0,0,0.6);
    border: 1px solid var(--border);
  }}

  header {{
    display:flex;
    align-items:center;
    justify-content:space-between;
    gap:16px;
    margin-bottom:18px;
  }}

  h1 {{
    margin:0;
    color:var(--title);
    font-weight:800;
    font-size:28px;
    letter-spacing:0.2px;
  }}

  .time {{
    color:var(--subtitle);
    font-weight:800; /* bold and bright per request */
    font-size:16px;
    background:transparent;
  }}

  .controls {{
    display:flex;
    gap:12px;
    align-items:center;
  }}

  .btn {{
    background:transparent;
    border:1px solid rgba(255,255,255,0.06);
    color:var(--muted);
    padding:8px 12px;
    border-radius:8px;
    cursor:pointer;
    font-weight:600;
    transition: all 160ms ease;
  }}
  .btn:hover{{
    transform:translateY(-2px);
    box-shadow: 0 6px 20px rgba(10,132,166,0.12);
    color:var(--title);
    border-color:rgba(0,229,255,0.12);
  }}

  /* Table */
  .candidate-table {{
    width:100%;
    border-collapse:collapse;
    overflow:hidden;
    border-radius:8px;
    table-layout:fixed;
  }}

  .candidate-table thead th {{
    text-align:left;
    padding:14px 16px;
    font-weight:700;
    color:var(--title);
    font-size:16px;
    position:sticky;
    top:0;
    background:linear-gradient(180deg, rgba(255,255,255,0.01), rgba(255,255,255,0.00));
    border-bottom:1px solid var(--border);
  }}

  .candidate-table tbody td {{
    padding:12px 16px;
    font-size:14px;
    color:var(--muted);
    white-space:nowrap;
    overflow:hidden;
    text-overflow:ellipsis;
  }}

  .candidate-table tbody tr.odd {{
    background:var(--row-odd);
  }}
  .candidate-table tbody tr.even {{
    background:var(--row-even);
  }}

  /* subtle name highlight cell, not too contrasty */
  .candidate-table .name-cell {{
    background:var(--name-highlight);
    color:var(--muted);
    border-radius:4px;
    padding:10px 14px;
  }}

  /* responsive */
  @media (max-width:900px) {{
    .candidate-table thead th:nth-child(2),
    .candidate-table tbody td:nth-child(2) {{ display:none; }} /* hide Name on small screens */
  }}

  .footer {{
    margin-top:16px;
    display:flex;
    justify-content:space-between;
    gap:8px;
    color:rgba(255,255,255,0.25);
    font-size:13px;
  }}
</style>

<!-- Auto-refresh the page every 3 seconds (same as before) -->
<script>
  setTimeout(() => {{
    window.location.reload();
  }}, 3000);
</script>
</head>
<body>
  <div class="container">
    <header>
      <h1>🎓 KTech Candidate Records</h1>
      <div style="display:flex;flex-direction:column;align-items:flex-end;">
        <div class="time">{now}</div>
        <div style="height:6px;"></div>
        <div class="controls">
          <button class="btn" onclick="location.reload()">Refresh</button>
          <button class="btn" onclick="window.print()">Print</button>
        </div>
      </div>
    </header>

    <main>
      {table_html}
    </main>

    <div class="footer">
      <div>Showing records from <strong style="color:var(--title)">{os.path.basename(EXCEL_PATH)}</strong></div>
      <div>Auto-refresh every 3s • Dark tech theme</div>
    </div>
  </div>
</body>
</html>
'''
    return render_template_string(page)

if __name__ == '__main__':
    # Use port 5000 or 80 as you prefer. 80 may need elevated privileges.
    app.run(host='0.0.0.0', port=5000, debug=True)
