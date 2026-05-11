NeuroToxCom Seeker V1.0 & Associated Visualization Scripts
NeuroToxCom Seeker V1.0 is a professional software tool for neurotoxic compound data analysis, integrating toxicology databases, mass spectrum matching, and analytical tools.
 Key Features
- **LC-MS Mode**: For liquid chromatography-mass spectrometry data analysis
- **GC-MS Mode**: For gas chromatography-mass spectrometry data analysis
 Basic Workflow
1. Select analysis method (LC-MS or GC-MS)
2. Select experimental data file (.mzML format)
3. Select spectral library file (.msp format)
4. Configure analysis parameters
5. Click "Start Analysis" to run
6. Results are automatically saved to the original file folder
 Batch Processing
- Supports processing entire folders
- Each file generates an independent result file
- Naming format: [original_filename]_[method]_results.xlsx
 Software Directory Structure
Software/
├── data/                            Data directory (required)
│   ├── all.xlsx                     Main database
│   ├── Physcicochemical.xlsx        Physicochemical property data
│   ├── *.msp                        Spectral library files
│   └── *.Help.docx                  Help documentation
└── NeuroToxCom_Seeker.py           Main program

Visualization Scripts (Data provided in Drawing data folder)
All scripts generate publication-ready vector graphics with transparent backgrounds, Arial fonts, and compact sizes (5-10 cm) in SVG/PDF formats.
 1. PBMT-Chord Diagram.py
Reads PBMT classification data from 1.PBMT-Chord Diagram date.xlsx, generates a chord diagram showing only specified node pairs (12 relationships, e.g., P-B, P-M), exports connection statistics to Excel, and outputs a 9×9 cm chord diagram.
 2. Physicochemical Ternary Plot.py
Reads physicochemical parameters (log KOA, log KAW, TE, etc.) from Physicochemical Ternary Plot date.xlsx, plots a ternary distribution map, and identifies/annotates compounds based on PBT and PMT criteria.
 3. Logkow -HLB-Classy.py
Reads log Kow and HLB data, creates horizontal bar charts for all compounds and each chemical class showing value ranges. Highlights partition boundaries, mean, and median with red dashed lines and colored markers. Linear scale on left, log scale on right.
 4. AOP.py
Reads three Excel files (AOP abbreviation mapping, compound-AOP associations, AOP event associations), constructs an interactive network graph (blue AOP event nodes, green compound nodes) with draggable nodes and S-key saving functionality.
 5. PBMT-Fitting Plot.py
Reads Physicochemical Ternary Plot date.xlsx, performs linear regression analysis on Log KOA vs Log KAW for all compounds, PBT, PMT, and both. Generates four side-by-side scatter plots with regression lines, R² values, and p-values. Exports statistics to Excel.
 6. PBMT classy.py
Reads SMILES structures, computes Morgan fingerprints, reduces dimensionality via PCA and t-SNE. Visualizes molecules in 2D space with different colors for PBT/PBMT/PMT hazard classes and different shapes for compound superclasses. Outputs scatter plots and detailed statistical reports.
 7. t-SNE+HOMO-LUMO gap.py
Reads log Kow, HOMO-LUMO gap, and BertzCT data. Creates scatter plots (log Kow vs HOMO-LUMO gap) where point size represents BertzCT (molecular complexity). Colors distinguish PBT (red) and PMT (blue). Generates two 8×7 cm SVG vector graphics.
 8. Structure Alert.py
Reads co-occurrence data of structural alert fragments, builds a network graph where node size represents fragment frequency and edge thickness represents co-occurrence rate. Supports draggable node layout and S-key saving to SVG/PDF/PNG.
 9. Retrospective - Concentration Violin Plot.py
Reads compound intensity data, groups by chemical class (classes with <20 samples merged into "Others"), performs Kruskal-Wallis + Dunn's test for significance. Plots log-scale violin plots sorted by median. Outputs statistical tables and 6×7 cm vector graphics.
 10. Retrospective - Compound Type Intensities Across Countries.py
Reads compound intensity data by category across countries. Generates choropleth world maps for each category, with colors filled by country code. Outputs transparent-background, unannotated SVG maps.
 11. Toxpi.py
Reads ToxPi scores and source data. Creates scatter plots ranked by descending score, with different colors for different sources (Literature search/Industrial list/Both). Outputs 9×5 cm transparent-background PDF/SVG vector graphics.
 12. MASST - Data Sources Across Countries.py
Reads compound counts by data source across countries. Creates horizontal stacked bar charts (log scale) sorted by total count, with total labels at the end of each row. Outputs 7×7 cm transparent-background SVG vector graphics.
 13. MASST-Bubble Plot.py
Reads compound counts across different sources. Generates bubble plots for each chemical class (x-axis: data sources, y-axis: compound names), where bubble size represents value and color represents source type. Also outputs a separate SVG legend file containing both color and size legends.
 Repository Structure
GitHub Repository/
├── NeuroToxCom_Seeker.py            Main program
├── data/                            Data directory
│   ├── all.xlsx
│   ├── Physcicochemical.xlsx
│   ├── *.msp
│   └── *.Help.docx
├── Drawing data/                    Input data for visualization scripts
│   ├── 1.PBMT-Chord Diagram date.xlsx
│   ├── Physicochemical Ternary Plot date.xlsx
│   └── ... (other data files)
├── PBMT-Chord Diagram.py
├── Physicochemical Ternary Plot.py
├── Logkow -HLB-Classy.py
├── AOP.py
├── PBMT-Fitting Plot.py
├── PBMT classy.py
├── t-SNE+HOMO-LUMO gap.py
├── Structure Alert.py
├── Retrospective - Concentration Violin Plot.py
├── Retrospective - Compound Type Intensities Across Countries.py
├── Toxpi.py
├── MASST - Data Sources Across Countries.py
└── MASST-Bubble Plot.py
Total: 14 Python scripts with corresponding data files uploaded to GitHub.
