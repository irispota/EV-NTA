# Specifies directory for raw files
data.files.dir <- "add file path here"

# Specifies directory to output the sample.list.file
data.dir <- "add file path here"

# Names the sample list output file
sample.list.file <- "sample.list.xlsx"

# Load required packages for data manipulation and visualization
# If these packages are not installed, run "install.packages()" before.
library(dplyr)
library(ggplot2)
library(writexl)
library(viridis)
library(readxl)
library(tidyverse)

# Sets the working directory to the location containing the raw files
setwd(data.files.dir)

# Reads a list of the file names in the current working directory
file.list <- list.files()

# Creates a data frame to store sample information
# Source and attributes are initially empty, will be filled from user input
sample.list <- data.frame(
  filename = file.list,
  source = NA,
  attribute1 = NA,
  attribute2 = NA
)

# Sets the working directory and outputs the sample.list.file
setwd(data.dir)
write_xlsx(sample.list, sample.list.file)

# Prompts user to fill in the generated Excel sheet with sample metadata before continuing
stop("Fill generated Excel Sheet with Sample Metadata.")

# Reads the filled-in sample file with attribute information
sample.list <- read_xlsx(sample.list.file)

# Sets working directory to access raw files
setwd(data.files.dir)

# Reads the content of the first file to initialize a data frame of the correct length
file.content <- readLines(file.list[1], warn = FALSE)

## Handles potential non-unicode characters and convert comma-separated numbers
file.content <- iconv(file.content, "UTF-8", "ASCII", sub = "")
file.content <- gsub(",", ".", file.content)

# Finds the line containing "Median Volume" (indicates start of relevant data)
start.index <- grep("Median Volume", file.content)

# Creates a data frame containing only the relevant data from the start index + 3 lines (skipping entire header)
df1 <- read.table(
  text = file.content[(start.index + 3) : length(file.content)], 
  header = FALSE
)

colnames(df1) <- c("Size", "Number", "Concentration", "Volume", "Area")

# Finds the line containing "-1" in the "Size" column (indicates end of relevant data)
end.index <- min(which(df1$Size == -1))

# The new data frame contains only relevant data to initialize a functional data frame
concentration.data <- df1[1:(end.index - 1), ]
concentration.data <- concentration.data[, 1:2]

# This function adds all relevant data and multiplies particle counts with the dilution factor
# The data is added to the concentration.data data frame
process.data.file <- function(filename) {
  file.content <- readLines(filename, warn = FALSE)
  file.content <- iconv(file.content, "UTF-8", "ASCII", sub = "")
  file.content <- gsub(",", ".", file.content)
  
  dilution.path <- grep("Dilution::", file.content, value = TRUE)
  dilution.factor <- as.numeric(gsub(".*Dilution::\\s*(\\d+\\.?\\d*).*", "\\1", dilution.path))
  
  start.index <- grep("Median Volume", file.content)
  
  preliminary.concentration.data <- read.table(
    text = file.content[(start.index + 3):length(file.content)], 
    header = FALSE,
    col.names = c("Size", "Number", "Concentration", "Volume", "Area")
  )
  
  end.index <- min(which(preliminary.concentration.data$Size == -1))
  preliminary.concentration.data <- preliminary.concentration.data[1:(end.index - 1), ]
  preliminary.concentration.data <- mutate_all(preliminary.concentration.data, as.numeric)
  preliminary.concentration.data$Concentration <- preliminary.concentration.data$Concentration * dilution.factor
  concentration.data <- cbind(concentration.data, preliminary.concentration.data$Concentration)
  
  return(concentration.data)
}

# Runs the function on each file within the raw files directory
for (filename in file.list) {
  concentration.data <- process.data.file(filename)
}

# Updates the data frame to remove the second column (used only for initialization) and renaming
concentration.data <- concentration.data[, -2]
colnames(concentration.data) <- c("Size", file.list)

# Transforms all values into numeric
concentration.data <- mutate_all(concentration.data, as.numeric)

# Calculates the sum of all values in each column (= total particle counts)
# concentration.data.total.particle: new data frame containing also total particle counts
concentration.data.total.particle <- rbind(colSums(concentration.data), concentration.data)

# Renames the first row
concentration.data.total.particle[1, 1] <- "Total Particles"

# Creates a new folder and writes an Excel file
dir.create("Total Particle Counts")
write_xlsx(concentration.data.total.particle, "Total Particle Counts/Total Particle counts.xlsx")

# Groups by attribute 1 and 2 and creates a subset data frame
sample.list.subset <- sample.list %>%
  group_by(attribute1, attribute2) %>%
  summarise(filename = list(filename), source = list(source),
            .groups = "drop")

# This function generates the plots for samples with matching attributes
## matching.cols: matches data from concentration.data with the sample.list.subset grouped based on attributes
## concentration.data.filter: stores filtered data to be used in plot
## df3: stores all values under the size of 800 nm in a numeric format
## plot: contains the plot, aesthetics: viridis
## dir.create: creates a new folder
## filename.plot and ggsave: generate filename for the plot based on sample attributes and save the plot
generate.replicate.plot <- function(concentration.data, sample.list.subset, index) {
  df3 <- concentration.data[, 1:2]
  matching.cols <- intersect(sample.list.subset$filename[[index]], names(concentration.data))
  concentration.data.filter <- concentration.data %>%
    select(all_of(matching.cols))
  
  df3 <- cbind(df3[, -2], concentration.data.filter)
  colnames(df3) <- c("Size", sample.list.subset$source[[index]])
  df3.filter <- filter(df3, Size <= 800)
  df3.long <- df3.filter %>%
    pivot_longer(cols = -Size, names_to = "Donor", values_to = "Particles/mL")
  df3.long$Donor <- factor(df3.long$Donor, levels = sort(unique(df3.long$Donor)))
  
  plot <- ggplot(df3.long, aes(x = Size, y = `Particles/mL`, color = Donor)) +
    geom_line(size = 1) +
    labs(title = paste0(sample.list.subset$attribute1[index], "-", sample.list.subset$attribute2[index]),
         x = "Size",
         y = "Particles/mL",
         color = "Donor") +
    scale_x_continuous(breaks = seq(0, max(df3.long$Size), by = 100)) +
    scale_color_viridis_d(labels = levels(df3.long$Donor)) +
    theme_minimal(base_size = 20) +
    theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5))
  
  dir.create("Compiled Graphs")
  filename.plot <- paste0(sample.list.subset$attribute1[index], "-", sample.list.subset$attribute2[index], ".png")
  ggsave(paste0("Compiled Graphs/", filename.plot), plot = plot, width = 20, height = 12, units = "cm", dpi = 600)
}

# applies the generate.replicate.plot function on each file in the raw file directory
for (i in 1:nrow(sample.list.subset)) {
  generate.replicate.plot(concentration.data, sample.list.subset, i)
}

# Additional plot with shading to indicate mean and standard error with a line and shading
generate.grouped.plot <- function(concentration.data, sample.list.subset) {
  for (i in 1:nrow(sample.list.subset)) {
    df3 <- concentration.data[, 1:2]
    matching.cols <- intersect(sample.list.subset$filename[[i]], names(concentration.data))
    concentration.data.filter <- concentration.data %>%
      select(all_of(matching.cols))
    
    df3 <- cbind(df3[, -2], concentration.data.filter)
    colnames(df3) <- c("Size", unlist(sample.list.subset$source[[i]]))
    df3.filter <- filter(df3, Size <= 800)
    df3.long <- df3.filter %>%
      pivot_longer(cols = -Size, names_to = "Donor", values_to = "Particles/mL")
    df3.long$Donor <- factor(df3.long$Donor, levels = sort(unique(df3.long$Donor)))
    
    # Calculate mean and standard error
    summarized_data <- df3.long %>%
      group_by(Size) %>%
      summarise(
        mean_particles = mean(`Particles/mL`),
        se_particles = sd(`Particles/mL`) / sqrt(n()),
        .groups = "drop"
      )
    
    plot <- ggplot(summarized_data, aes(x = Size, y = mean_particles)) +
      geom_ribbon(aes(ymin = mean_particles - se_particles, 
                      ymax = mean_particles + se_particles), 
                  fill = "#7e94c4", alpha = 0.5) +
      geom_line(size = 1, color = "#7e94c4") +
      labs(title = paste0(sample.list.subset$attribute1[i], " - ", sample.list.subset$attribute2[i]),
           x = "Size",
           y = "Particles/mL") +
      scale_x_continuous(breaks = seq(0, max(summarized_data$Size), by = 100)) +
      theme_minimal(base_size = 20) +
      theme(axis.text.x = element_text(angle = 90, hjust = 1, vjust = 0.5))
    
    dir.create("Compiled Graphs", showWarnings = FALSE)
    filename.plot <- paste0("Combined_Plot_", sample.list.subset$attribute1[i], "_", sample.list.subset$attribute2[i], ".pdf")
    ggsave(paste0("Compiled Graphs/", filename.plot), plot = plot, width = 20, height = 12, units = "cm", dpi = 600)
  }
}

# Apply the generate.grouped.plot function
generate.grouped.plot(concentration.data, sample.list.subset)
