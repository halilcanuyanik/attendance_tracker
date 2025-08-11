# Attendance Tracking System

A desktop-based attendance management application built with **Python**, **Tkinter**, **Pandas**, and **OpenPyXL**.  
This system allows staff members to check in and out using their ID numbers (scanned or manually entered), while maintaining a log in an Excel file.  
It also highlights unregistered staff entries for easy identification.

---

## 📌 Features

- **Barcode/ID Entry Support**  
  Accepts scanned or manually typed 11-digit ID numbers.
  
- **Automated Entry/Exit Logging**  
  Automatically detects if an employee is entering or exiting based on the last recorded action for that day.
  
- **Excel-based Storage**  
  All attendance records are stored in `attendance_log.xlsx`, making them easy to share and manage.
  
- **Highlighting Unregistered Staff**  
  Entries with unregistered staff IDs are automatically marked in red for follow-up.
  
- **User-friendly Interface**  
  Built with Tkinter for an intuitive and simple design, with real-time clock display.
  
- **Auto-clear Information**  
  Recorded details are shown briefly and then cleared automatically after 5 seconds.

---

## 📂 Project Structure
- ├── staff_list.xlsx # Contains registered staff details (ID, First Name, Last Name, Department)
- ├── attendance_log.xlsx # Automatically created/updated attendance log
- ├── attendance_tracker.py # Main Python script

---


## 📊 Excel File Details

### staff_list.xlsx  
Must include:
| ID Number | First Name | Last Name | Department |
|-----------|------------|-----------|------------|
| 12345678901 | John      | Doe       | IT         |

### attendance_log.xlsx  
Created/updated automatically:
| Date       | ID Number   | First Name | Last Name | Department | Entry Time | Exit Time |
|------------|-------------|------------|-----------|------------|------------|-----------|
| 2025-08-11 | 12345678901 | John       | Doe       | IT         | 09:00      | 17:00     |

Unregistered staff (`X X X`) will be highlighted in **red**.

---

## 🚀 Installation & Usage

### 1️⃣ Install Requirements

```bash
pip install pandas openpyxl
```

### 2️⃣ Prepare Staff List
Ensure staff_list.xlsx exists in the project folder with correct columns.

### 3️⃣ Run the Application
```bash
python attendance_tracker.py
```
### 4️⃣ Record Attendance
- Scan an ID card or type the 11-digit ID and press Enter or click Submit.
- The system will log an Entry or Exit depending on the previous record.

### ⚠️ Notes
- Only exactly 11-digit numeric IDs are accepted.
- Unregistered IDs will still be logged but highlighted in red.
- Make sure to close attendance_log.xlsx before recording attendance, or the save will fail.

---

## 🛠 Technologies Used
- **Python3**
- **Tkinter** - GUI Interface
- **Pandas** - Data handling
- **OpenPyXL** – Excel read/write operations

---

## License
This project is licensed under the MIT License.
