# NeuroToxCom
Complete code for a neurotoxic emerging pollutant compound database and visualization system. This system includes data management, analysis, and visualization functions.
1. Software Introduction
NeuroToxCom Seeker V1.0 is a professional neurotoxicological compound data analysis software that integrates toxicology databases, mass spectrometry matching, and analytical tools.
2. Main Functional Modules
2.1 Main Interface - Compound Search
•	Search Methods:
Single Search: Enter CAS number or compound name
Batch Search: Upload TXT file (one CAS per line)
•	Operation Buttons:
🔍 Search: Execute search
🧪 Batch Search: Upload file for batch query
🗑️ Clear: Restore all data
⬇️ Download: Export search results
2.3 Mass Spectrometry Matching Tool
Supports two analysis modes:
•	LC-MS Mode: For liquid chromatography-mass spectrometry data analysis
•	GC-MS Mode: For gas chromatography-mass spectrometry data analysis
Basic Workflow:
1.	Select analysis method (LC-MS or GC-MS)
2.	Select experimental data file (.mzML format)
3.	Select spectral library file (.msp format)
4.	Set analysis parameters
5.	Click "START ANALYSIS" to begin
6.	Results automatically saved to original file folder
Batch Processing:
•	Supports processing entire folders
•	Each file generates separate results
•	Naming format: [original_filename]_[method]_results.xlsx
3. Spectral Library Management
Available Spectral Libraries:
•	All Spectra (81,585 spectra)
•	LC-MS Spectra (70,135 spectra)
•	LC-ESI-QTOF (22,758 spectra)
•	LC-MSMS Positive/Negative Mode
•	GC-MS Spectra (17,111 spectra)
•	GC-EI Spectra (9,759 spectra)
Usage Method:
1.	Click "Spectral Library" in sidebar
2.	Click "⬇ Download" button in table
3.	Select save location to download
4. File Requirements
4.1 Supported File Formats
•	Data Files: .xlsx, .csv, .txt
•	Mass Spectrometry Files: .mzML (experimental data), .msp (spectral library)
•	Document Files: .docx (help documents)
