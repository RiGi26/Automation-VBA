🚀 Advanced VBA Excel Automation Engine
A high-performance Excel-based automation system engineered with VBA (Visual Basic for Applications). This project transforms standard spreadsheets into an intelligent data processing application featuring anti-freeze optimization, dynamic UI components, and self-cleaning database logic.

📌 Project Overview
This system was developed to solve common enterprise spreadsheet issues such as performance degradation, data redundancy, and manual formatting errors. It features an event-driven architecture that responds to user inputs in real-time while maintaining application stability.

🚩 Key Problems Solved

Performance Lag: Prevents Excel from freezing during large data operations.


Data Integrity: Eliminates manual VLOOKUP errors through automated logic.


Maintenance Overhead: Automates database hygiene by removing obsolete entries.


UI Inconsistency: Standardizes professional reporting styles automatically.

💡 Technical Features
1. Intelligent Anti-Freeze Logic
To maintain high responsiveness, the engine includes a pre-execution filter that detects and bypasses automation during massive data changes (e.g., deleting >50 rows or multiple columns at once).

VBA

' Example: Anti-Freeze Solution 
If Target.Rows.Count > 50 Or Target.Columns.Count > 5 Then
    Application.StatusBar = "⚠️ Large change detected - Auto-process skipped to avoid freeze"
    Exit Sub
End If
2. Dynamic Cascading Dropdowns
Utilizes custom functions like CreateDropdownCDE to inject Data Validation into cells dynamically based on previous selections (Column B → C → D → E), ensuring a guided and error-free user journey.

3. Self-Cleaning Database Management
The CleanupDATABASE routine intelligently monitors data usage. If a specific entry is no longer referenced in the main TEMPLATE sheet, it is automatically purged from the DATABASE to keep the file lightweight and efficient.

4. Pattern-Based Visual Highlighting
Implements a smart highlighting engine (HighlightKolomBU) that uses pattern matching (Regex-like logic) to color-code rows based on specific input values (e.g., 6-digit numeric codes or specific alphanumeric prefixes).

5. Pixel-Perfect Auto-Formatting

Dynamic Bordering: Automatically applies thin borders only to active data rows.


Automated Cell Merging: Merges and unmerges cells dynamically for cleaner report aesthetics without breaking data integrity.

🛠️ Technical Stack

Language: Excel VBA.


Logic: Event-Driven Programming (Worksheet_Change), Alphanumeric Generators, Pattern Matching.


Architecture: XML-based Workbook management with Custom UI Controls.


📂 Installation & Usage
Clone this repository.

Open the .xlsm file.

Ensure Macros are enabled.

Use the DisableVBA and EnableVBA buttons if you need to perform bulk manual pastes.

👨‍💻 Developed By
Riyadh Akhdan Syafi Data Automation Specialist | Google Apps Script & Python Developer 📧 [riyadhakhdan3@gmail.com]
