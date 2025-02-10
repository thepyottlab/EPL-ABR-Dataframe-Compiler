# EPL ABR Dataframe Compiler
<p align="center">
  <img src="https://github.com/user-attachments/assets/80dcc902-ebc6-473c-95bd-41ea616f36b1" width="600" height="283">
</p>

## Description
EPL ABR Dataframe Compiler takes a folder with analyzed EPL Auditory Brainstem Response (ABR) .txt files (data files can be present in the same folder without issue) that were generated with [ABR Peak Analysis](https://github.com/EPL-Engineering/abr-peak-analysis) (version 1.10.1). Based on P and N amplitudes, ABR wave 1-5 are calculated for each frequency and are displayed in a 'long' dataframe format for data analysis. Amplitudes that were measured at intensities under the set threshold are automatically set to 0. The merged, analyzed files are exported in a single .xlsx file with color-coded values for analysis.

## Features
- Searches for all .txt files in a folder.
- Combines analyzed .txt files to a single dataframe in a 'long' format and exports to a .xslx file.
- Exports ID, frequency, threshold, threshold method, stimulus intensity, correlation coefficient, average noise, noise standard deviation, N1-5 & P1-5 amplitudes, N1-5 & P1-5 latencies, and wave 1-5 amplitudes.
- Automatically calculates ABR wave 1-5 amplitudes
- Compatible with having 'Do noise floor analysis' ticker or unticked.
- Color codes the cells from light to dark depending on the value of the cells for easy visual identification of outlayers or strange values.
- Filtered columns automatically turned on.

## Installation
1. Go to the [Releases](https://github.com/thepyottlab/EPL-ABR-Dataframe-Compiler/releases) page.
2. Download the `EPL_ABR_Dataframe_Compiler_Setup.exe` file.
3. Run the installer and follow the on-screen instructions.

## Usage
1. Run the EPL ABR Dataframe Compiler.
2. Select the folder that contains the analyzed ABR files and press 'Select folder'.
3. Wait for the compilation of the dataframe to complete.
4. A file called 'Merged_dataframe.xlsx' should be created in the selected folder.
