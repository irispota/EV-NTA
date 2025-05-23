# EV-NTA
## Documentation for NTA Data Processing and Plotting Script
Author: Iris Potakowskyj
Date: 2025-05-23

Please cite: 10.5281/zenodo.15497706

### Description:

This script processes and analyzes data from Nanoparticle Tracking Analysis (NTA) experimentscommonly used for extracellular vesicle size and concentration measurments. 
It is compatible with data from ZetaView (Particle Metrix). To use this script for other NTA instruments, modify indices to fit data structure.

It performs the following tasks:
1. Reads data files: The script reads data files in .txt format from a specified directory.
2. Generates an Excel file: This excel file contains the sample information. 
3. Calculates particle concentrations: It calculates the particle concentration based on the data in the files.
4. Generates plots: The script generates plots of particle concentration vs. particle size for different experimental conditions.
5. Creates summary tables: It creates a summary table with total particle counts for each sample.

### Test Files:
The test files provided are from qEVoriginal columns, pooled fractions 1-5 from 2 different donors.

### Input:
* data.dir: The directory where the sample.list.file will be saved. data.dir has to be different than data.files.dir.
* data.files.dir: The directory containing the raw data files. The folder can only contain the text files to be analyzed.
* sample.list.file: An Excel file containing information about the samples, such as source and attributes. This file will be generated in the first steps of the code and has to be filled in manually with the sample information and saved. !! If it is not completed, the code will potentially not be able to group the files appropriately. 
	Fill in as needed: use "source" to specify, e.g. the donor or cell line. 
	Use "attribute1" and "attribute2" to specify other variables that you want to use for grouping or file labelling, e.g. fractions, treatment conditions, isolation method.
	The headers of the columns need to remain unchanged!

### Output:
* Excel file: A file containing the total particle counts for each sample.
* PNG images: Plots of particle concentration vs. particle size for different experimental conditions.

### Running Environment:
* The code is open within R Studio and run from the control panel.
* The "Run Source" (Ctrl+Shift+L) can be used to run the entire code. To fill in the sample.list.file the code will be paused automatically.

### Dependencies:
* R packages: `dplyr`, `ggplot2`, `writexl`, `viridis`, `readxl`, `tidyverse`

### Script Steps:

1. Set working directory: Manually set the working directories for "data.dir" and "data.files.dir".
2. Choose name for sample list: Fill in a custom name for the "sample.list.file".
3. The sample.list.file will be ceated. After this step has been completed, a stop function with the text "Fill generated Excel Sheet with Sample Metadata." will be shown in the console. The user has to manually navigate to the data.dir directory and fill in the sample information as described above. 
4. Process individual data files: 
    * The script iterates through each data file.
    * For each file, it extracts the dilution factor and reads the particle size and concentration data.
    * It calculates the final concentration by multiplying the raw concentration with the dilution factor.
5. Combine data from all files: The script combines the data from all processed files into a single data frame.
6. Calculate total particle counts: It calculates the total particle count for each size range and adds a row with the total counts to the data frame.
7. Create summary table: The script writes the data frame containing total particle counts to an Excel file.
8. Generate plots:
    * The script groups samples based on their attributes in the sample list.
    * For each group, it generates a plot of particle concentration vs. particle size for each sample within the group.
    * The plots are saved as PNG images with meaningful file names.


### Handling error messages:

* The script includes a warning message indicating that the directory "Compiled Graphs" might already exist. This message can be ignored as the script handles the case where the directory already exists.
* The script cannot handle empty experiment files (text files not containing particle data), before running the code, these files should be removed.
* The script includes a section to correct numbers that use "," instead of "." as a decimal separator, other variations in data format needed to be handled manually.

