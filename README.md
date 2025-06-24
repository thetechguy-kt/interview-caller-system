# <b>Interview Queue Caller System</b>
# <i>A simple, offline interview management system with a queue-based caller system, built using Python and Tkinter.</i>

This is a multi-part application designed to streamline the candidate flow management during interviews or walk-in assessments. It automates token generation, queue control, and live displays for interview counters. Built with Python and Tkinter, the system uses Excel and JSON files for data storage and operates entirely offline.

## 📄 1. Candidate Token Generator App - `Candidate POS.py (With Packaged .exe File for Windows)`
This is the front-desk application where a staff member logs each candidate as they arrive. It:
  - Allows entry of candidate name and contact number
  - Automatically assigns and displays a daily token number
  - Saves each entry into an Excel file `candidate_list.xlsx`
  - Generates a printable PDF ticket for the candidate with QR code and interview info
  - Resets the token count every day automatically
  - Stores token data in a daily folder under `Tickets/YYYY-MM-DD - Tickets`
  - Tracks date using `config/last_ticket_date.txt`

> <b> Ideal for reception or registration desk staff to quickly log and print token slips. </b>

## 🖥️ 2. Interview Room Control Panel - `Interview Room 1.py, Interview Room 2.py (With Packaged .exe File for Windows for both the Python Files)`
This app runs inside each interview room and allows the interviewer or assistant to call the next candidate. Each instance represents one counter/room and includes:
  - A main control window (e.g., for Room 1)
  - A pop-up display window visible to candidates
  - A Call Next button that selects the next available token
  - Recall, Waiting, Open/Close Room controls
  - Updates a central file queue_state.json with the list of called tokens
  - Only reads from `candidate_list.xlsx` (does not write to it)

> <b> Multiple rooms can run their own instance (Room 1, Room 2, and more), and all will coordinate via the shared `queue_state.json`. </b>

## 📺 3. Central Display Board - `Central Display.py (With Packaged .exe File for Windows)`
This is the live token display screen visible to waiting candidates. It auto-refreshes every few seconds and shows:
  - The current token number and candidate name
  - The room number where the candidate should go
  - A clean layout suitable for large screens or TV monitors
It pulls data from:
  - `queue_state.json` → Called token data (Updated by Room apps)
  - `candidate_list.xlsx` → Candidate names and details

> <b>This app is read-only and does not modify any files. Place it in the same folder as the shared `.json` and `.xlsx` files for live updates.</b>

## 🗂️ 4. Record Viewer App - `Record Viewer.py (With Packaged .exe File for Windows)`
This utility app is designed for admins or HR staff to monitor, review, or audit the list of registered or interviewed candidates. It:
  - Loads and displays the contents of `candidate_list.xlsx` in a tabular view
  - Uses a simple Tkinter-based table/grid or a widget like `ttk.Treeview`
  - Supports scrolling for large datasets
  - Auto-refreshes every 3 seconds to reflect new candidate entries
  - Is fully read-only — it does not modify the Excel file

> <b>This app is especially helpful during busy interview sessions for non-technical users who need a live, automatically refreshing view of which candidates have registered and when. It’s also ideal for verifying past entries, performing audit checks, or exporting data manually — all without needing to open Excel separately.</b>

# 📁 File Overview
| File/Folder | App/File Name | Description |
| :---: | :---: | --- |
| `Candidate POS.py` | Token Generator App	 | Registers candidates, assigns daily token numbers, and generates printable PDF tickets with QR codes. |
| `Interview Room 1\2`</center> | Interview Room Controller | Calls the next candidate, updates `queue_state.json`, and displays the token info in-room. |
| `Central Display.py` | Central Display Board | Displays currently called tokens and assigned rooms in real-time for waiting candidates. |
| `Record Viewer.py` | Live Record Viewer App | Shows and auto-refreshes the full list of registered candidates from `candidate_list.xlsx`. |
| `candidate_list.xlsx` | Excel File - Candidate List | Stores all logged candidate details including name, number, time, and assigned token. |
| `queue_state.json` | JSON File - Queue State | Maintains the live state of called tokens and their assigned interview rooms. |
| `config/last_ticket_date.txt` | Text File - Last Ticket Date | Tracks the last active date for auto-resetting token numbers each day. |
| `Tickets/YYYY-MM-DD - Tickets/` | PDF File - Tokens, Excel File - List of Tokens | Daily folder containing all generated PDF tickets plus a copy of the Excel log `candidate_list_YYYY-MM-DD.xlsx` for that day. |

<b> Note: 
  - Place all files in a single folder. Also include `dip_config/notify.wav`, which plays a sound and highlights the name when a new candidate is called from Room 1, 2, etc. You can     change the sound path in the Python File. Compile using PyInstaller or similar to create a `.exe File`. In the `.exe Files` the sound is embedded within the app.
  - `.exe Files` can be accessed from by Clicking Here: [Drive Share Folder](https://drive.google.com/drive/folders/1dhkN6V82qp-A2-ePw5LvCwpez8Fb67YU?usp=sharing)
</b>

