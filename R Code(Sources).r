# Load necessary libraries
library(readxl)
library(dplyr)
library(ggplot2)
library(reshape2)

# Load the Excel file (adjust the path as necessary)
file_path <- "Source_Comparison.xlsx"

# Read each sheet into a list while excluding the "Directory" sheet
sheets <- excel_sheets(file_path)
data_list <- lapply(sheets, function(sheet) {
  if (sheet != "Directory") {
    read_excel(file_path, sheet = sheet, col_names = FALSE) %>%
      pull(1) %>%
      na.omit()
  } else {
    NULL
  }
})

# Remove NULL entries from the list
data_list <- data_list[!sapply(data_list, is.null)]
sheets <- sheets[!sapply(data_list, is.null)]

# Get all unique terms across all sheets
all_terms <- unique(unlist(data_list))

# Initialize a matrix for repeated terms
result_matrix <- matrix(0, nrow = length(all_terms), ncol = length(sheets),
                        dimnames = list(all_terms, sheets))

# Fill the matrix
for (i in seq_along(data_list)) {
  for (term in data_list[[i]]) {
    if (term %in% all_terms) {
      result_matrix[term, sheets[i]] <- 1
    }
  }
}

# Convert the matrix to a data frame for better viewing
result_df <- as.data.frame(result_matrix)
result_df <- cbind(Term = rownames(result_df), result_df)
rownames(result_df) <- NULL

# Create a long format data frame for ggplot
long_df <- melt(result_df, id.vars = "Term")

# Create a color mapping for 1 and 0
long_df$Color <- ifelse(long_df$value == 1, "blue", "white")

# Calculate the number of terms and split into three parts
n_terms <- nrow(result_df)
terms_per_page <- ceiling(n_terms / 3)

# Create the PDF and save the plots
pdf("matrix_output.pdf", width = 10, height = 8)

# Loop to create plots for each page
for (page in 0:2) {
  start_index <- page * terms_per_page + 1
  end_index <- min((page + 1) * terms_per_page, n_terms)
  
  # Subset the long_df for the current page
  page_terms <- result_df$Term[start_index:end_index]
  page_data <- long_df[long_df$Term %in% page_terms, ]

  # Create the plot for the current page
  p <- ggplot(page_data, aes(x = variable, y = Term)) +
    geom_tile(aes(fill = Color), color = "black") +
    scale_fill_identity() +
    labs(title = paste("Matrix of Repeated Terms Across Sheets - Page", page + 1)) +
    theme_minimal() +
    theme(axis.text.x = element_text(angle = 45, hjust = 1))

  print(p)  # Print the plot to the PDF
}

# Close the PDF device
dev.off()


Heat map: Displays the frequency of shared terms or data points across sheets

Version 1:

# Load necessary libraries
library(readxl)
library(dplyr)
library(ggplot2)
library(reshape2)

# Load the Excel file (adjust the path as necessary)
file_path <- "Source_Comparison.xlsx"

# Read each sheet into a list while excluding the "Directory" sheet
sheets <- excel_sheets(file_path)
data_list <- lapply(sheets, function(sheet) {
  if (sheet != "Directory") {
    read_excel(file_path, sheet = sheet, col_names = FALSE) %>%
      pull(1) %>%
      na.omit()  # Remove NA values
  } else {
    NULL
  }
})

# Remove NULL entries from the list
data_list <- data_list[!sapply(data_list, is.null)]
sheets <- sheets[!sapply(data_list, is.null)]

# Create a matrix to store counts of shared terms
n <- length(data_list)
shared_matrix <- matrix(0, nrow = n, ncol = n, dimnames = list(sheets, sheets))

# Calculate shared terms
for (i in 1:n) {
  for (j in 1:n) {
    if (i != j) {
      shared <- length(intersect(data_list[[i]], data_list[[j]]))
      shared_matrix[i, j] <- shared
    }
  }
}

# Convert matrix to data frame for ggplot
shared_df <- as.data.frame(as.table(shared_matrix))

# Create the heatmap
ggplot(shared_df, aes(Var1, Var2, fill = Freq)) +
  geom_tile() +
  scale_fill_gradient(low = "white", high = "blue") +
  labs(title = "Heatmap of Shared Terms Between Sheets",
       x = "Sheet",
       y = "Sheet",
       fill = "Shared Terms") +
  theme_minimal()


Word cloud:

# Load necessary libraries
library(readxl)
library(dplyr)
library(wordcloud)  # For creating word clouds
library(RColorBrewer)  # For color palettes

# Load the Excel file (adjust the path as necessary)
file_path <- "Source_Comparison_1.xlsx"

# Read each sheet into a list while excluding the "Directory" sheet
sheets <- excel_sheets(file_path)
data_list <- lapply(sheets, function(sheet) {
  if (sheet != "Directory") {  # Exclude the "Directory" sheet
    read_excel(file_path, sheet = sheet, col_names = FALSE) %>%
      pull(1) %>%
      na.omit()  # Remove NA values
  } else {
    NULL  # Return NULL for the "Directory" sheet
  }
})

# Remove NULL entries from the list
data_list <- data_list[!sapply(data_list, is.null)]

# Create a single character vector of all sources
all_sources <- unlist(data_list)

# Create a data frame to count occurrences of each unique value
word_embeddings <- as.data.frame(table(all_sources))

# Rename the columns for clarity
colnames(word_embeddings) <- c("Word", "Count")

# Convert counts to numeric
word_embeddings$Count <- as.numeric(word_embeddings$Count)

# Set the JPEG output
jpeg("Word_Cloud.jpeg", width = 800, height = 600)

# Generate the word cloud
wordcloud(words = word_embeddings$Word, 
          freq = word_embeddings$Count, 
          min.freq = 1, 
          max.words = 100, 
          random.order = FALSE, 
          rot.per = 0.35, 
          scale = c(4, 0.5), 
          colors = brewer.pal(8, "Dark2"))

# Close the JPEG device
dev.off()

# Set the PDF output
pdf("wordcloud1.pdf", width = 10, height = 8)

# Generate the word cloud again for PDF
wordcloud(words = word_embeddings$Word, 
          freq = word_embeddings$Count, 
          min.freq = 1, 
          max.words = 100, 
          random.order = FALSE, 
          rot.per = 0.35, 
          scale = c(4, 0.5), 
          colors = brewer.pal(8, "Dark2"))

# Close the PDF device
dev.off()

cat("Word cloud saved as Word_Cloud.jpeg and wordcloud1.pdf\n")




Network graph: network graph that visualizes the relationships between different sheets based on shared data sources

# Load necessary libraries
library(readxl)
library(dplyr)
library(igraph)

# Load the Excel file (adjust the path as necessary)
file_path <- "Source Comparison 1.xlsx"

# Read each sheet into a list while excluding the "Directory" sheet
sheets <- excel_sheets(file_path)
data_list <- lapply(sheets, function(sheet) {
  if (sheet != "Directory") {  # Exclude the "Directory" sheet
    read_excel(file_path, sheet = sheet, col_names = FALSE) %>%
      pull(1) %>%
      na.omit()  # Remove NA values
  } else {
    NULL  # Return NULL for the "Directory" sheet
  }
})

# Remove NULL entries from the list
data_list <- data_list[!sapply(data_list, is.null)]
sheets <- sheets[!sapply(data_list, is.null)]  # Keep the corresponding sheet names

# Create a data frame to store shared terms between sheets
shared_terms <- data.frame(sheet1 = character(), sheet2 = character(), shared_count = integer(), stringsAsFactors = FALSE)

# Find shared terms between each pair of sheets
for (i in 1:(length(data_list) - 1)) {
  for (j in (i + 1):length(data_list)) {
    shared <- intersect(data_list[[i]], data_list[[j]])
    if (length(shared) > 0) {
      shared_terms <- rbind(shared_terms, data.frame(sheet1 = sheets[i], 
                                                     sheet2 = sheets[j], 
                                                     shared_count = length(shared)))
    }
  }
}

# Check if there are any shared terms
if (nrow(shared_terms) == 0) {
  stop("No shared terms found between sheets.")
}

# Create a graph from the shared_terms data frame
graph <- graph_from_data_frame(shared_terms, directed = FALSE)

# Set node colors
V(graph)$color <- rainbow(length(V(graph)))  # Assign rainbow colors to nodes

# Set edge colors based on shared_count
edge_colors <- colorRampPalette(c("lightblue", "darkblue"))(max(shared_terms$shared_count))
E(graph)$color <- edge_colors[shared_terms$shared_count]

# Set PDF output
pdf("Network graph 1.pdf", width = 10, height = 8)

# Plot the network graph
plot(graph, 
     vertex.label = V(graph)$name, 
     vertex.size = 5, 
     edge.width = E(graph)$shared_count, 
     edge.arrow.size = 0.5, 
     layout = layout_with_fr,
     edge.color = E(graph)$color)  # Use edge colors

# Close the PDF device
dev.off()

cat("Network graph saved as Network graph 1.pdf\n")



UpSet plot:

# Install and load necessary packages
# If you don't have these packages installed, run these lines first:
# install.packages("UpSetR")
# install.packages("readr") # Often a dependency, good to include
# install.packages("readxl") # Package for reading Excel files

library(UpSetR)
library(readr)
library(readxl) # Load the readxl package

# Define the path to your Excel file
# IMPORTANT: Ensure this file is in your R working directory,
# or provide the full path to the file (e.g., "C:/Users/YourUser/Documents/Source Comparison 1.xlsx").
excel_file_path <- "Source_Comparison.xlsx"

# Define the names of the sheets you want to compare from the Excel file
# >>> VERY IMPORTANT: Ensure these names exactly match your Excel sheet tabs <<<
sheet_names <- c(
  "Symmap",
  "ETCM 1.0",
  "ETCM 2.0",
  "HERB 1.0",
  "HERB 2.0",
  "TM-MC 2.0",
  "TCMBank",
  "TCM-ID 2.0",
  "DCABM-TCM",
  "CMAUP 2.0",
  "ITCM",
  "BATMAN TCM" # <<< CHANGED: Now "BATMAN TCM" as requested >>>
)

# Create an empty list to store the unique elements from each sheet as sets
list_of_sets <- list()

message("--- Starting to read data from Excel sheets ---")

# Loop through each sheet name to read its content and prepare the data for UpSetR
for (sheet_name in sheet_names) {
  message(paste0("  Processing sheet: '", sheet_name, "'"))
  
  # Read the specific sheet from the Excel file.
  # We assume no header and that the data is in the first column.
  # `read_excel` from `readxl` is used for robust reading.
  data <- tryCatch({
    read_excel(excel_file_path, sheet = sheet_name, col_names = FALSE, .name_repair = "minimal")
  }, error = function(e) {
    message(paste("  ERROR: Could not read sheet '", sheet_name, "' from file '", excel_file_path, "' - ", e$message, sep = ""))
    message("  ACTION REQUIRED: Please ensure the sheet name is exact and the file is accessible/unlocked.")
    return(NULL) # Return NULL if there's an error
  })

  # Check if data was read successfully and is not empty
  if (!is.null(data) && nrow(data) > 0) {
    # Extract the first column (which contains the items), convert to character,
    # trim leading/trailing whitespace, and get unique values.
    # Trimming whitespace is crucial for accurate intersection detection.
    current_set <- unique(trimws(as.character(data[[1]])))
    list_of_sets[[sheet_name]] <- current_set
    message(paste0("  Successfully read '", sheet_name, "'. Found ", length(current_set), " unique items."))
  } else {
    # If data is NULL (due to error) or has 0 rows
    message(paste("  WARNING: Sheet '", sheet_name, "' in '", excel_file_path, "' is empty or could not be read. Adding as an empty set.", sep = ""))
    list_of_sets[[sheet_name]] <- character(0) # Add an empty set if sheet is empty/unreadable
  }
}

message("--- Finished reading data from Excel sheets ---")

# --- Debugging: Print structure of list_of_sets before plotting ---
message("\n--- Summary of sets prepared for UpSetR ---")
non_empty_sets_count <- 0
for (name in names(list_of_sets)) {
  set_size <- length(list_of_sets[[name]])
  message(paste0("  Set '", name, "': ", set_size, " unique items."))
  if (set_size > 0) {
    non_empty_sets_count <- non_empty_sets_count + 1
  }
}
message(paste0("Total non-empty sets prepared: ", non_empty_sets_count, " out of ", length(list_of_sets), " requested sets."))
message("--- End of set summary ---\n")


# --- Generate the UpSet Plot and Save as PDF ---

# Set the output PDF file name and dimensions.
# !!! Increased width and height even further as requested !!!
# Be aware that very large dimensions can result in large PDF file sizes.
pdf_output_filename <- "Source_UpSet_Plot.pdf"
pdf(pdf_output_filename, width = 30, height = 20) # Significantly increased dimensions

message(paste0("Generating UpSet plot and saving to '", pdf_output_filename, "'..."))

# Generate the UpSet plot
# Key parameters:
# - fromList(list_of_sets): Converts the list of unique elements into the format UpSetR expects.
# - sets = names(list_of_sets): Explicitly tells UpSetR to consider all named sets.
# - order.by = "freq", decreasing = TRUE: Orders the intersections by their frequency (size), largest first.
# - nintersects = NA: THIS IS CRUCIAL. It tells UpSetR to display ALL possible non-empty intersections.
#                     If the plot still appears visually overwhelming, it's because there are
#                     many actual intersections. You could then limit this, e.g., 'nintersects = 50'
#                     or set a 'min_size' for intersections to display.
# - text.scale: Adjusts the size of various text elements on the plot for readability.
# - mainbar.y.label, sets.x.label: Labels for axes.
# - point.size, line.size: Visual adjustments for the intersection matrix.
# - mb.ratio: Adjusts the ratio of the main bar plot to the set size bar plot.
upset(fromList(list_of_sets),
      sets = names(list_of_sets),
      order.by = "freq",
      decreasing = TRUE,
      nintersects = NA, # Display all non-empty intersections
      text.scale = c(1.5, 1.5, 1.2, 1.2, 1.5, 1),
      mainbar.y.label = "Intersection Size",
      sets.x.label = "Set Size",
      point.size = 3,
      line.size = 1,
      mb.ratio = c(0.7, 0.3)
)

# Close the PDF device. This is crucial for the file to be written.
dev.off()

message(paste0("UpSet plot saved as '", pdf_output_filename, "' in your current working directory."))
message("--- UpSet plot generation complete ---")

Bar Charts Sources:

# Load necessary libraries
library(readxl)
library(dplyr)
library(ggplot2) # For creating the bar chart

# Load the Excel file (adjust the path as necessary)
file_path <- "Source Comparison 1.xlsx"

# Read each sheet into a list while excluding the "Directory" sheet
sheets <- excel_sheets(file_path)
data_list <- lapply(sheets, function(sheet) {
  if (sheet != "Directory") { # Exclude the "Directory" sheet
    read_excel(file_path, sheet = sheet, col_names = FALSE) %>%
      pull(1) %>%
      na.omit() # Remove NA values
  } else {
    NULL # Return NULL for the "Directory" sheet
  }
})

# Remove NULL entries from the list
data_list <- data_list[!sapply(data_list, is.null)]
# sheets <- sheets[!sapply(data_list, is.null)] # Not strictly needed for this specific task, but good practice

# --- New code for bar chart of repeated cells ---

# 1. Combine all data into a single vector
all_elements <- unlist(data_list)

# 2. Count the frequency of each unique element
# Using table() for counts, then converting to a data frame
element_counts <- as.data.frame(table(all_elements), stringsAsFactors = FALSE)
colnames(element_counts) <- c("Element", "Frequency")

# 3. Order the elements by frequency for better visualization
element_counts <- element_counts %>%
  arrange(desc(Frequency))

# 4. Create the bar chart
# Filter for elements that appear more than once (optional, depends on what 'repeated' means to you)
# Here, we'll plot all, but you can add filter(Frequency > 1) if needed.
ggplot(element_counts, aes(x = reorder(Element, -Frequency), y = Frequency)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  labs(
    title = "Frequency of Elements Across All Sheets",
    x = "Element",
    y = "Number of Occurrences"
  ) +
  theme_minimal() +
  theme(axis.text.x = element_text(angle = 45, hjust = 1)) # Rotate x-axis labels if they overlap

# You can save this plot to a PDF or other image format
# ggsave("Commonly_Repeated_Cells_Bar_Chart.pdf", width = 10, height = 7)
# ggsave("Commonly_Repeated_Cells_Bar_Chart.png", width = 10, height = 7)

# --- End of new code ---

# The original UpSet plot generation code (optional, you can keep it or remove if only bar chart is needed)
# sheets <- sheets[!sapply(data_list, is.null)] # Re-enable if you removed it for UpSetR
# venn_list <- setNames(data_list, sheets)
# binary_df <- data.frame(matrix(0, nrow = length(unique(unlist(data_list))), ncol = length(data_list)))
# colnames(binary_df) <- sheets
# rownames(binary_df) <- unique(unlist(data_list))
# for (i in seq_along(data_list)) {
#   binary_df[unique(data_list[[i]]), i] <- 1
# }
# upset(as.data.frame(binary_df), sets = colnames(binary_df), keep.order = TRUE)

UpSet Plot analysing headers of databases:

# Load necessary libraries
library(readxl)
library(dplyr)
library(tidyr)  # Load tidyr for pivot_longer
library(UpSetR)
library(grid)   # Load grid for text placement

# Read the Excel file
data <- read_excel("Headers.xlsx", sheet = "Sheet1")

# View the structure of the data to understand its layout
print(head(data))

# Convert data to a long format for easier processing
data_long <- data %>%
  pivot_longer(cols = everything(), names_to = "Column", values_to = "Terms") %>%
  filter(!is.na(Terms) & Terms != "")

# Create a list of unique terms for each column
term_list <- data_long %>%
  group_by(Column) %>%
  summarise(Terms = list(unique(Terms)))

# Create a binary presence-absence data frame
term_matrix <- as.data.frame(matrix(0, nrow = length(unique(data_long$Terms)), ncol = length(unique(data_long$Column))))
colnames(term_matrix) <- unique(data_long$Column)
rownames(term_matrix) <- unique(data_long$Terms)

# Fill the matrix with 1s for presence of terms
for (i in unique(data_long$Column)) {
  term_matrix[data_long$Terms[data_long$Column == i], i] <- 1
}

# Save UpSet Plot as a PDF
pdf("Summarised UpSet Plot.pdf", width = 10, height = 7)

# Generate UpSet Plot
upset(as.data.frame(term_matrix), sets = colnames(term_matrix), 
      main.bar.color = "blue", sets.bar.color = "orange",
      keep.order = TRUE, order.by = "freq")

# Close the PDF device to finalize the plot
dev.off()

# Reopen the PDF to add text labels
pdf("Summarised UpSet Plot.pdf", width = 10, height = 7)

# Generate UpSet Plot again for labeling
upset(as.data.frame(term_matrix), sets = colnames(term_matrix), 
      main.bar.color = "blue", sets.bar.color = "orange",
      keep.order = TRUE, order.by = "freq")

# Add labels to the bars using the grid package
grid.text(colnames(term_matrix), 
          x = seq(0.1, 0.9, length.out = length(colnames(term_matrix))), 
          y = unit(1, "npc") * 1.05, 
          gp = gpar(col = "black", fontsize = 10), 
          just = "center")

# Close the PDF device
dev.off()



Bar Chart with Table: For information


# Load necessary libraries
library(readxl)
library(dplyr)    # For data manipulation (e.g., piping)
library(ggplot2)  # For creating the bar chart

# Load the Excel file
excel_file <- "Reference.xlsx"

# --- IMPORTANT: Ensure the Excel file is in your working directory ---
# You can check your current working directory with getwd()
# You can set your working directory with setwd("path/to/your/directory")
# For example, if "Reference 1.xlsx" is on your Desktop:
# setwd("C:/Users/YourUsername/Desktop") # For Windows
# setwd("/Users/YourUsername/Desktop")   # For macOS/Linux

if (!file.exists(excel_file)) {
  stop(paste0("Error: The file '", excel_file, "' was not found in your current working directory.\n",
              "Please ensure the file is present and the working directory is set correctly.\n",
              "Current working directory: ", getwd()))
}

# Get sheet names
sheet_names <- excel_sheets(excel_file)

# Initialize a data frame to store the results
results <- data.frame(Sheet = character(), FilledCells = integer(), stringsAsFactors = FALSE)

# Loop through each sheet and count filled cells
cat("--- Counting Filled Cells in Each Sheet ---\n")
for (sheet in sheet_names) {
  cat(paste0("Processing sheet: '", sheet, "'...\n"))
  data <- read_excel(excel_file, sheet = sheet)
  filled_cells <- sum(!is.na(data))  # Count non-NA values
  results <- rbind(results, data.frame(Sheet = sheet, FilledCells = filled_cells))
  cat(paste0("  Sheet '", sheet, "': ", filled_cells, " filled cells.\n"))
}
cat("--------------------------------------------\n\n")

# Prepare data for bar chart
# Ensure Sheet names are factors and ordered for consistent plotting
results$Sheet <- factor(results$Sheet, levels = results$Sheet[order(results$FilledCells, decreasing = TRUE)])


# Create the bar chart
bar_chart_pdf_name <- "Bar Chart - Filled Cells.pdf"
pdf(bar_chart_pdf_name, width = 10, height = 7) # Open PDF device

ggplot(results, aes(x = Sheet, y = FilledCells)) +
  geom_bar(stat = "identity", fill = "steelblue", color = "black") + # stat="identity" uses y-value directly
  labs(
    title = "Number of Filled Cells in Each Sheet",
    x = "Sheet Name",
    y = "Number of Filled Cells"
  ) +
  theme_minimal() +
  theme(
    axis.text.x = element_text(angle = 45, hjust = 1, size = 10), # Rotate x-axis labels for readability
    axis.text.y = element_text(size = 10),
    axis.title = element_text(size = 12, face = "bold"),
    plot.title = element_text(hjust = 0.5, size = 16, face = "bold"),
    panel.grid.major.x = element_blank(), # Remove vertical grid lines
    panel.grid.minor.x = element_blank()
  ) +
  geom_text(aes(label = FilledCells), vjust = -0.5, size = 3.5) # Add labels on top of bars

dev.off() # Close PDF device

cat(paste("Bar chart of filled cells saved to:", bar_chart_pdf_name, "\n"))

Radar Chart for Data in Sources:

# Load necessary libraries
library(readxl)
library(fmsb)

# Load the Excel file
excel_file <- "Reference.xlsx"

# Get sheet names
sheet_names <- excel_sheets(excel_file)

# Initialize a data frame to store the results
results <- data.frame(Sheet = character(), FilledCells = integer(), stringsAsFactors = FALSE)

# Loop through each sheet and count filled cells
for (sheet in sheet_names) {
  data <- read_excel(excel_file, sheet = sheet)
  filled_cells <- sum(!is.na(data))  # Count non-NA values
  results <- rbind(results, data.frame(Sheet = sheet, FilledCells = filled_cells))
}

# Prepare data for radar chart
max_cells <- max(results$FilledCells)
radar_data <- as.data.frame(t(results$FilledCells))
colnames(radar_data) <- results$Sheet
radar_data <- rbind(rep(max_cells, ncol(radar_data)), rep(0, ncol(radar_data)), radar_data)

# Create the radar chart
pdf("Radar Chart.pdf", width = 8, height = 6)  # Open PDF device
radarchart(radar_data, 
           axistype = 1,
           pcol = "blue", 
           pfcol = scales::alpha("blue", 0.3), 
           plwd = 2, 
           plty = 1,
           title = "Number of Filled Cells in Each Sheet",
           axislabcol = "grey",
           cglcol = "grey",
           cglty = 1,
           cglwd = 0.8)
dev.off()  # Close PDF device
