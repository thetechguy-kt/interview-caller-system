from flask import Flask, render_template_string
from openpyxl import load_workbook

app = Flask(__name__)

def excel_to_html(filepath):
    wb = load_workbook(filepath, data_only=True)
    ws = wb.active

    rows = list(ws.rows)
    if not rows:
        return "<p>No data found in Excel.</p>"

    html = '<table border="1" cellpadding="5" cellspacing="0" style="border-collapse: collapse; width: 100%;">'

    # Header
    html += '<thead><tr>'
    for cell in rows[0]:
        html += f'<th>{cell.value if cell.value is not None else ""}</th>'
    html += '</tr></thead>'

    # Body
    html += '<tbody>'
    for row in rows[1:]:
        html += '<tr>'
        for cell in row:
            html += f'<td>{cell.value if cell.value is not None else ""}</td>'
        html += '</tr>'
    html += '</tbody></table>'

    return html

@app.route('/')
def index():
    table_html = excel_to_html('candidate_list.xlsx')
    page = f'''
    <!DOCTYPE html>
    <html>
    <head>
        <title>IITC Candidate Record Viewer</title>
        <style>
            body {{ font-family: Arial, sans-serif; margin: 20px; }}
            table {{ border-collapse: collapse; width: 100%; }}
            th, td {{ border: 1px solid #ccc; padding: 8px; text-align: left; }}
            thead {{ background-color: #f2f2f2; }}
        </style>
        <script>
            // Reload entire page every 3 seconds
            setTimeout(() => {{
                window.location.reload();
            }}, 3000);
        </script>
    </head>
    <body>
        <h2>Candidate List</h2>
        {table_html}
    </body>
    </html>
    '''
    return render_template_string(page)

if __name__ == '__main__':
    app.run(debug=True)
