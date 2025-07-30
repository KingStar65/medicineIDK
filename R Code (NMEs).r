#Code for converting common chemical names to InChIKey format:

# Install and load necessary packages
# install.packages(c("openxlsx", "httr", "jsonlite")) # Uncomment to install if not already installed
library(openxlsx)
library(httr)
library(jsonlite)

# Function to convert a chemical name (common or IUPAC) to InChIKey using PubChem PUG-REST with retries and throttling
convert_name_to_inchikey <- function(chemical_name, max_retries = 3, initial_delay = 0.5) {
  # Handle empty or non-character inputs gracefully
  if (is.na(chemical_name) || !is.character(chemical_name) || nchar(trimws(chemical_name)) == 0) {
    return(NA)
  }

  encoded_name <- URLencode(chemical_name, reserved = TRUE)
  inchikey_url <- paste0("https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/", encoded_name, "/property/InChIKey/JSON")

  inchikey_result <- NA
  
  for (attempt in 1:max_retries) {
    response <- tryCatch({
      GET(inchikey_url, timeout(20)) # Increased timeout slightly for complex queries
    }, error = function(e) {
      message(paste("Attempt", attempt, ": Error fetching InChIKey for", chemical_name, ":", e$message))
      return(NULL)
    })

    if (!is.null(response) && http_status(response)$category == "Success") {
      json_content <- content(response, "text", encoding = "UTF-8")
      if (json_content != "") { # Check for empty content
        json_data <- fromJSON(json_content)
        # Check if PropertyTable and Properties exist and contain InChIKey
        if (!is.null(json_data$PropertyTable) &&
            !is.null(json_data$PropertyTable$Properties) &&
            "InChIKey" %in% names(json_data$PropertyTable$Properties)) {
          inchikey_result <- json_data$PropertyTable$Properties$InChIKey[1]
          # If successful, break out of retry loop
          return(inchikey_result)
        }
      }
    }
    
    # If not successful, wait before retrying (exponential backoff)
    if (attempt < max_retries) {
      delay <- initial_delay * (2^(attempt - 1)) # 0.5s, 1s, 2s, etc.
      message(paste("Retrying in", delay, "seconds for", chemical_name))
      Sys.sleep(delay)
    }
  }
  
  # If all retries fail
  message(paste("Failed to retrieve InChIKey after", max_retries, "attempts for:", chemical_name))
  return(NA)
}

# --- Main Script ---

# 1. Determine the path to the Desktop dynamically
desktop_path <- file.path(Sys.getenv("HOME"), "Desktop")
# For some Windows configurations, Sys.getenv("USERPROFILE") might be more reliable
# desktop_path <- file.path(Sys.getenv("USERPROFILE"), "Desktop")


# Define your input and output Excel file details
# The input and output file is the SAME file: "IUPAC.xlsx" on your Desktop
input_file_name <- "IUPAC.xlsx"
excel_file_path <- file.path(desktop_path, input_file_name) # This is now the single path for read/write
sheet_name <- "Sheet1" # Assuming your chemical names are in Sheet1

# 2. Read chemical names from the Excel file (Column A, header 'IUPAC_Name')
tryCatch({
  message(paste("Attempting to read from:", excel_file_path))
  existing_data <- read.xlsx(excel_file_path, sheet = sheet_name)

  if (!"IUPAC_Name" %in% colnames(existing_data)) {
    stop("Column 'IUPAC_Name' not found in the Excel file. Please ensure Column A is named 'IUPAC_Name'.")
  }

  chemical_names_from_file <- existing_data$IUPAC_Name # Read from the column that previously held IUPAC names

  message(paste("Read", length(chemical_names_from_file), "chemical names from", excel_file_path))

  # 3. Convert chemical names to InChIKeys
  message("Starting InChIKey conversion via PubChem API. This may take some time due to retries and throttling...")
  
  inchikey_results <- character(length(chemical_names_from_file)) # Pre-allocate
  for (i in seq_along(chemical_names_from_file)) {
    current_name <- chemical_names_from_file[i]
    message(paste0("Processing compound ", i, "/", length(chemical_names_from_file), ": ", current_name))
    inchikey_results[i] <- convert_name_to_inchikey(current_name)
    
    # Explicit throttling: Wait to respect PubChem's rate limits
    Sys.sleep(0.25)
  }
  
  message("Finished InChIKey conversion.")

  # 4. Add the new "InChIKey" column to the existing data frame
  # This will populate Column B if it's the next available column, or update an existing 'InChIKey' column.
  existing_data$InChIKey <- inchikey_results

  # --- Debugging Step: Check data before writing ---
  message("Preview of data about to be written:")
  print(head(existing_data))
  message("End of preview.")
  # --- End Debugging Step ---

  # 5. Save the updated data frame back to the SAME Excel file and SAME sheet
  write.xlsx(existing_data, excel_file_path, sheetName = sheet_name, rowNames = FALSE, overwrite = TRUE)

  message(paste("Successfully processed and updated results in:", excel_file_path))
  message("The InChIKeys have been written to Column B.")

}, error = function(e) {
  message(paste("An error occurred:", e$message))
  message(paste("Please ensure the file '", input_file_name, "' is directly on your Desktop and has a column 'IUPAC_Name' in Sheet1.", sep=""))
  message("Also, verify that the Excel file is not open while the script is running.")
})



Code for converting IUPAC names to InChIKey:

# Install and load necessary packages
# install.packages(c("openxlsx", "httr", "jsonlite")) # Uncomment to install if not already installed
library(openxlsx) # Using openxlsx for robust Excel reading/writing
library(httr)     # For making API requests
library(jsonlite) # For parsing JSON responses

# Function to convert IUPAC name to InChIKey using PubChem PUG-REST
convert_iupac_to_inchikey <- function(iupac_name) {
  # Handle empty or non-character inputs gracefully
  if (is.na(iupac_name) || !is.character(iupac_name) || nchar(trimws(iupac_name)) == 0) {
    return(NA)
  }

  # Sanitize IUPAC name for URL (e.g., replace spaces with %20)
  encoded_iupac <- URLencode(iupac_name, reserved = TRUE)

  # PubChem PUG-REST URL to get InChIKey from IUPAC name
  inchikey_url <- paste0("https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/name/", encoded_iupac, "/property/InChIKey/JSON")

  inchikey_result <- NA # Initialize result to NA

  response <- tryCatch({
    # Add a timeout to prevent infinite waits for unresponsive APIs
    GET(inchikey_url, timeout(10)) # 10-second timeout
  }, error = function(e) {
    message(paste("Error fetching InChIKey for", iupac_name, ":", e$message))
    return(NULL)
  })

  if (!is.null(response) && http_status(response)$category == "Success") {
    json_content <- content(response, "text", encoding = "UTF-8")
    if (json_content != "") { # Check for empty content
      json_data <- fromJSON(json_content)
      # Check if PropertyTable and Properties exist and contain InChIKey
      if (!is.null(json_data$PropertyTable) &&
          !is.null(json_data$PropertyTable$Properties) &&
          "InChIKey" %in% names(json_data$PropertyTable$Properties)) {
        inchikey_result <- json_data$PropertyTable$Properties$InChIKey[1]
      }
    }
  }

  return(inchikey_result)
}

# --- Main Script ---

# 1. Determine the path to the Desktop dynamically
desktop_path <- file.path(Sys.getenv("HOME"), "Desktop")
# For some Windows configurations, Sys.getenv("USERPROFILE") might be more reliable
# desktop_path <- file.path(Sys.getenv("USERPROFILE"), "Desktop")


# Define your input and output Excel file details
# The input and output file is the SAME file: "IUPAC.xlsx" on your Desktop
input_file_name <- "IUPAC.xlsx"
excel_file_path <- file.path(desktop_path, input_file_name) # This is now the single path for read/write
sheet_name <- "Sheet1" # Assuming your IUPAC names are in Sheet1

# 2. Read IUPAC names from the Excel file (Column A)
tryCatch({
  message(paste("Attempting to read from:", excel_file_path))
  # Read the entire sheet content
  existing_data <- read.xlsx(excel_file_path, sheet = sheet_name)

  # Ensure the column exists and is named "IUPAC_Name" (case-sensitive)
  if (!"IUPAC_Name" %in% colnames(existing_data)) {
    stop("Column 'IUPAC_Name' not found in the Excel file. Please ensure Column A is named 'IUPAC_Name'.")
  }

  iupac_names_from_file <- existing_data$IUPAC_Name

  message(paste("Read", length(iupac_names_from_file), "IUPAC names from", excel_file_path))

  # 3. Convert IUPAC names to InChIKeys
  message("Starting InChIKey conversion via PubChem API. This may take some time...")
  inchikey_results <- sapply(iupac_names_from_file, convert_iupac_to_inchikey, USE.NAMES = FALSE)
  message("Finished InChIKey conversion.")

  # 4. Add the new "InChIKey" column to the existing data frame
  # This ensures Column A (IUPAC_Name) remains untouched, and Column B is added/updated.
  existing_data$InChIKey <- inchikey_results

  # --- Debugging Step: Check data before writing ---
  message("Preview of data about to be written:")
  print(head(existing_data))
  message("End of preview.")
  # --- End Debugging Step ---

  # 5. Save the updated data frame back to the SAME Excel file and SAME sheet
  # write.xlsx from openxlsx will overwrite the specified sheet's content
  write.xlsx(existing_data, excel_file_path, sheetName = sheet_name, rowNames = FALSE, overwrite = TRUE)

  message(paste("Successfully processed and updated results in:", excel_file_path))
  message("The InChIKeys have been written to Column B.")

}, error = function(e) {
  message(paste("An error occurred:", e$message))
  message(paste("Please ensure the file '", input_file_name, "' is directly on your Desktop and has a column 'IUPAC_Name' in Sheet1.", sep=""))
  message("Also, verify that the Excel file is not open while the script is running.")
})



Convert PubChem CID to common chemical names:

# Install and load necessary packages
# install.packages(c("openxlsx", "httr", "jsonlite")) # Uncomment to install if not already installed
library(openxlsx)
library(httr)     # For making API requests
library(jsonlite) # For parsing JSON responses

# Function to convert PubChem CID to common chemical name (Title) using PubChem PUG-REST
convert_cid_to_common_name <- function(cid_number, max_retries = 3, initial_delay = 0.5) {
  # Handle empty or non-numeric inputs gracefully
  if (is.na(cid_number) || !is.numeric(cid_number) || cid_number <= 0) {
    return(NA_character_) # Return NA_character_ for string column
  }

  # Ensure CID is an integer for the URL
  cid_number <- as.integer(cid_number)

  # PubChem API endpoint for CID to Property (Title)
  # The 'Title' property often gives the most common chemical name.
  # If you want ALL synonyms, you'd use '/synonyms/JSON' instead of '/property/Title/JSON'
  # and then parse the array of synonyms (e.g., take the first one or combine them).
  common_name_url <- paste0("https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound/cid/", cid_number, "/property/Title/JSON")

  common_name_result <- NA_character_ # Initialize with NA_character_ for string result
  
  for (attempt in 1:max_retries) {
    response <- tryCatch({
      # Increased timeout for complex queries or slow server responses
      GET(common_name_url, timeout(30)) 
    }, error = function(e) {
      message(paste("Attempt", attempt, ": Error fetching common name for CID", cid_number, ":", e$message))
      return(NULL)
    })

    if (!is.null(response) && http_status(response)$category == "Success") {
      json_content <- content(response, "text", encoding = "UTF-8")
      
      # Optional: Debugging - uncomment to see raw JSON response
      # message(paste("DEBUG: Raw JSON response for CID", cid_number, ":", json_content))

      if (json_content != "") { # Check for empty content
        json_data <- fromJSON(json_content)
        # Parse the JSON to extract the 'Title' property
        if (!is.null(json_data$PropertyTable) &&
            !is.null(json_data$PropertyTable$Properties) &&
            "Title" %in% names(json_data$PropertyTable$Properties)) {
          common_name_result <- json_data$PropertyTable$Properties$Title[1]
          # If successful, break out of retry loop
          return(common_name_result)
        }
      }
    }
    
    # If not successful, wait before retrying (exponential backoff)
    if (attempt < max_retries) {
      delay <- initial_delay * (2^(attempt - 1)) # 0.5s, 1s, 2s, etc.
      message(paste("Retrying in", delay, "seconds for CID", cid_number))
      Sys.sleep(delay)
    }
  }
  
  # If all retries fail
  message(paste("Failed to retrieve common name after", max_retries, "attempts for CID:", cid_number))
  return(NA_character_)
}

# --- Main Script ---

# 1. Determine the path to the Desktop dynamically
desktop_path <- file.path(Sys.getenv("HOME"), "Desktop")

# Define your input and output Excel file details
input_file_name <- "Code run 2.xlsx" # New file name
excel_file_path <- file.path(desktop_path, input_file_name) 
sheet_name <- "Sheet1" # Assuming your CIDs are in Sheet1

# --- Optional: Create a dummy input Excel file for demonstration ---
# This block is for demonstration only. If you have your own file, you can remove or comment it out.
if (!file.exists(excel_file_path)) {
  dir_path <- dirname(excel_file_path)
  if (!dir.exists(dir_path)) {
    dir.create(dir_path, recursive = TRUE)
    message(paste("Created directory:", dir_path))
  }

  dummy_data <- data.frame(
    CID = c(
      6857,    # Ethanol
      2244,    # Aspirin
      180,     # Acetone
      241,     # Benzene
      753,     # Glycerol
      3672,    # Ibuprofen
      962,     # Water
      2519,    # Caffeine
      5793,    # Glucose
      6013,    # Penicillin G
      99999999, # A very high, likely non-existent CID
      NA       # Blank row
    ),
    stringsAsFactors = FALSE
  )
  # Pre-populate column for Common_Name with NAs for the dummy file
  dummy_data$Common_Name <- NA_character_ 

  write.xlsx(dummy_data, excel_file_path, sheetName = sheet_name, row.names = FALSE)
  message(paste("Created a dummy input file at:", excel_file_path))
  message("Please check this file and add your PubChem CIDs to Column A (header 'CID').")
  message("Ensure you save the file before running the script again if you edit the dummy data.")
}
# --- End of Optional Dummy File Creation ---


# 2. Read PubChem CIDs from the Excel file (Column A, header 'CID')
tryCatch({
  message(paste("Attempting to read from:", excel_file_path))
  existing_data <- read.xlsx(excel_file_path, sheet = sheet_name)

  if (!"CID" %in% colnames(existing_data)) {
    stop("Column 'CID' not found in the Excel file. Please ensure Column A is named 'CID'.")
  }

  cid_numbers_from_file <- existing_data$CID # Read CIDs from the 'CID' column

  message(paste("Read", length(cid_numbers_from_file), "CIDs from", excel_file_path))

  # 3. Convert CIDs to common names
  message("Starting CID to common name conversion via PubChem API. This may take some time...")
  
  common_name_results <- character(length(cid_numbers_from_file)) # Pre-allocate for string results
  for (i in seq_along(cid_numbers_from_file)) {
    current_cid <- cid_numbers_from_file[i]
    message(paste0("Processing CID ", i, "/", length(cid_numbers_from_file), ": ", current_cid))
    common_name_results[i] <- convert_cid_to_common_name(current_cid)
    
    # Explicit throttling: Wait to respect PubChem's rate limits (0.5 seconds per query)
    Sys.sleep(0.5) 
  }
  
  message("Finished CID to common name conversion.")

  # 4. Add the new "Common_Name" column to the existing data frame
  # This will create a new column 'Common_Name' or overwrite an existing one.
  existing_data$Common_Name <- common_name_results

  # --- Debugging Step: Check data before writing ---
  message("Preview of data about to be written:")
  print(head(existing_data))
  message("End of preview.")
  # --- End Debugging Step ---

  # 5. Save the updated data frame back to the SAME Excel file and SAME sheet
  write.xlsx(existing_data, excel_file_path, sheetName = sheet_name, row.names = FALSE, overwrite = TRUE)

  message(paste("Successfully processed and updated results in:", excel_file_path))
  message("The common chemical names have been written to Column B.")

}, error = function(e) {
  message(paste("An error occurred:", e$message))
  message(paste("Please ensure the file '", input_file_name, "' is directly on your Desktop and has a column 'CID' in Sheet1.", sep=""))
  message("Also, verify that the Excel file is not open while the script is running.")
})






Compiled data:

# --- Install necessary packages if you haven't already ---
# Uncomment and run these lines if you get errors about missing packages.
# It's crucial that these installations complete successfully before running the rest of the script.
# install.packages("readxl")
# install.packages("dplyr")
# install.packages("proxy") # For Jaccard distance
# install.packages("ggplot2")
# install.packages("cluster") # For silhouette
# install.packages("factoextra") # For optimal k and clustering visualization (depends on ggpubr)
# install.packages("MASS") # For MDS (cmdscale)
# install.packages("ggpubr", dependencies = TRUE) # Explicitly install ggpubr and its dependencies
# install.packages("corrplot") # For visualizing the Jaccard coefficient matrix
# install.packages("gridExtra") # For printing tables to PDF
# install.packages("gtable") # Dependency for gridExtra table aesthetics

# Load the required libraries
library(readxl)
library(dplyr)
library(proxy)
library(ggplot2)
library(cluster)
library(factoextra) # This library requires 'ggpubr' for some of its plotting functions
library(MASS) # For cmdscale
library(stats) # For prcomp (PCA) and dist (Euclidean distance)
library(corrplot) # For visualizing the matrix
library(gridExtra) # For printing tables to PDF
library(gtable) # For grid.table aesthetics

# --- Configuration ---
input_file_name <- "Compiled_done.xlsx" # The Excel file generated in the previous step
output_column_names <- c("Herbs", "Formulas", "Targets", "Diseases", "Ingredients") # Expected column names

# --- Step 1: Read the Compiled_done.xlsx file and prepare data for analysis ---

# Check if the input file exists
if (!file.exists(input_file_name)) {
  stop(paste("Error: Input file '", input_file_name, "' not found. Please ensure it's in the same directory.", sep = ""))
}

# Get all sheet names from the compiled Excel file
compiled_sheets <- excel_sheets(input_file_name)

# If no sheets are found, stop the process
if (length(compiled_sheets) == 0) {
  stop(paste("No sheets found in '", input_file_name, "'. Cannot perform analysis.", sep = ""))
}

message(paste("Found ", length(compiled_sheets), " sheets in '", input_file_name, "'.", sep = ""))

# This list will store unique terms for each sheet (database) for Jaccard/PCA (all terms combined)
list_of_term_sets <- list()
# This will store unique terms counts for each sheet, broken down by column (for the summary table)
list_of_term_counts_per_column_per_sheet <- list()
# This will store all unique terms found across ALL sheets for the PCA matrix columns
all_unique_terms_global <- c()

# Loop through each sheet in the compiled Excel file
for (sheet_name in compiled_sheets) {
  message(paste("Processing sheet: '", sheet_name, "' for term extraction.", sep = ""))

  tryCatch({
    # Read the sheet without column names and skip the first row (headers)
    sheet_data_raw <- read_excel(input_file_name, sheet = sheet_name, col_names = FALSE, skip = 1)

    # Initialize a list to store counts for the current sheet's columns
    current_sheet_counts_by_column <- list()
    # Initialize a vector to hold all terms for the current sheet (combined for Jaccard/PCA)
    all_terms_in_sheet_combined <- c()

    # Iterate through the first 5 columns (corresponding to Herbs, Formulas, etc.) by index
    for (col_index in 1:length(output_column_names)) {
      col_name_expected <- output_column_names[col_index] # Get the expected column name (e.g., "Herbs")
      if (ncol(sheet_data_raw) >= col_index) { # Check if the column exists in the raw data
        # Extract terms from the current column, remove NAs, convert to character
        terms <- na.omit(as.character(sheet_data_raw[[col_index]]))
        # Convert terms to lowercase
        terms <- tolower(terms)
        unique_terms_col <- unique(terms) # Get unique terms for this specific column
        current_sheet_counts_by_column[[col_name_expected]] <- length(unique_terms_col) # Store count for summary
        all_terms_in_sheet_combined <- c(all_terms_in_sheet_combined, unique_terms_col) # Add to combined list
      } else {
        # If a column is missing in this specific sheet, store 0 for its count
        current_sheet_counts_by_column[[col_name_expected]] <- 0
        message(paste("Info: Column for '", col_name_expected, "' (index ", col_index, ") not found in sheet '", sheet_name, "'. Skipping term extraction for this column.", sep = ""))
      }
    }
    list_of_term_counts_per_column_per_sheet[[sheet_name]] <- current_sheet_counts_by_column
    list_of_term_sets[[sheet_name]] <- unique(all_terms_in_sheet_combined) # Store unique combined terms for Jaccard/PCA
    all_unique_terms_global <- c(all_unique_terms_global, unique(all_terms_in_sheet_combined)) # Add to global unique terms

  }, error = function(e) {
    message(paste("Error processing sheet '", sheet_name, "': ", e$message, sep = ""))
  })
}

# Remove any sheets that ended up with no terms (e.g., if they were completely empty or had no relevant columns)
list_of_term_sets <- list_of_term_sets[sapply(list_of_term_sets, length) > 0]
list_of_term_counts_per_column_per_sheet <- list_of_term_counts_per_column_per_sheet[names(list_of_term_sets)] # Keep only processed sheets

# Get all unique terms globally for the PCA matrix columns
all_unique_terms_global <- unique(all_unique_terms_global)

# If after processing, there are fewer than 2 databases with terms, analysis is not meaningful
if (length(list_of_term_sets) < 2) {
  stop("Not enough sheets with terms found to perform clustering or PCA (need at least 2).")
}

# --- Diagnostic: Print sample terms and Jaccard distances ---
message("\n--- Diagnostic Information ---")
message("Unique terms extracted for first few sheets (max 3):")
for (i in 1:min(3, length(list_of_term_sets))) {
  sheet_name <- names(list_of_term_sets)[i]
  terms <- list_of_term_sets[[sheet_name]]
  message(paste0("  '", sheet_name, "': ", paste(head(terms, 5), collapse = ", "), ifelse(length(terms) > 5, ", ...", "")))
}
message("----------------------------\n")

# --- Create Presence-Absence Matrix for PCA and Euclidean Distance ---
message("Creating presence-absence matrix for PCA and Euclidean Distance...")
# Initialize a matrix with zeros
presence_absence_matrix <- matrix(0,
                                  nrow = length(list_of_term_sets),
                                  ncol = length(all_unique_terms_global),
                                  dimnames = list(names(list_of_term_sets), all_unique_terms_global))

# Populate the matrix
for (sheet_name in names(list_of_term_sets)) {
  terms_in_sheet <- list_of_term_sets[[sheet_name]]
  # Set 1 for terms present in the current sheet
  presence_absence_matrix[sheet_name, terms_in_sheet] <- 1
}

# --- Step 2: Calculate Jaccard Distance & Euclidean Distance Matrices ---

message("Calculating Jaccard distance matrix (for reference and heatmap)...")
jaccard_distance_matrix <- proxy::dist(list_of_term_sets, method = "Jaccard")

message("\n--- Jaccard Distance Summary (for reference) ---")
jaccard_values <- as.vector(jaccard_distance_matrix)
print(summary(jaccard_values))
if (max(jaccard_values) < 0.01) {
  message("Warning: Maximum Jaccard distance is very low. This indicates high similarity between your databases based on Jaccard.")
  message("The dendrogram based on Jaccard would appear compressed due to this high similarity.")
}
message("-------------------------------\n")

# --- NEW: Euclidean Distance Matrix for Clustering and MDS ---
message("Calculating Euclidean distance matrix for clustering and MDS...")
# We use the 'dist' function from 'stats' package for Euclidean distance on the presence-absence matrix
euclidean_distance_matrix <- dist(presence_absence_matrix, method = "euclidean")

message("\n--- Euclidean Distance Summary ---")
euclidean_values <- as.vector(euclidean_distance_matrix)
print(summary(euclidean_values))
message("-------------------------------\n")


# --- NEW: Print Jaccard Coefficient (Similarity) Matrix to PDF ---
message("Generating Jaccard Similarity Matrix heatmap...")
jaccard_similarity_matrix <- as.matrix(1 - jaccard_distance_matrix)

pdf("Jaccard_Similarity_Matrix_Heatmap.pdf", width = 10, height = 10)
corrplot(jaccard_similarity_matrix,
         method = "color", # Use color to represent coefficients
         type = "full",    # Show full matrix
         order = "hclust", # Order by hierarchical clustering
         tl.col = "black", # Label color
         tl.srt = 45,      # Label text rotation
         addCoef.col = "black", # Add coefficients to the plot
         number.cex = 0.7, # Size of the coefficients
         main = "Jaccard Similarity Matrix Between Databases"
)
dev.off()
message("Jaccard Similarity Matrix heatmap saved as 'Jaccard_Similarity_Matrix_Heatmap.pdf'")


# --- Step 3: Hierarchical Clustering (using Euclidean Distance) ---

message("Performing Hierarchical Clustering (using Euclidean Distance)...")

# Perform hierarchical clustering using the Euclidean distance matrix
h_cluster_result <- hclust(euclidean_distance_matrix, method = "ward.D2") # Ward's method is common and effective

# Plot the dendrogram
# Save dendrogram to a PDF file for better quality
# Increased height for better spread, and adjusted ylim.
pdf("Hierarchical_Clustering_Dendrogram_Euclidean.pdf", width = 10, height = 10) # Changed filename
plot(h_cluster_result,
     main = "Hierarchical Clustering of Databases (Euclidean Distance)", # Changed title
     xlab = "Databases (Sheets)",
     ylab = "Distance",
     hang = -1, # hang = -1 aligns labels at the bottom
     ylim = c(0, max(h_cluster_result$height) * 1.05) # Set ylim to slightly above max height to show full range
     )
dev.off()
message("Hierarchical clustering dendrogram saved as 'Hierarchical_Clustering_Dendrogram_Euclidean.pdf'")


# --- Step 4: Determine Optimal K for K-means Clustering ---

message("Determining optimal K for K-means Clustering...")

# Max k is limited by the number of data points.
k_max <- min(length(compiled_sheets) - 1, 6) # Max k for silhouette, at most n-1 or 6
optimal_k <- 3 # Default optimal_k, will be updated if analysis is possible

if (k_max < 2) {
  message("Not enough data points for robust K-means or silhouette analysis (need at least 2). Defaulting to k=1.")
  optimal_k <- 1
} else {
  # For silhouette method, we need a data matrix. We can use the PCA or MDS coordinates.
  # Let's use the PCA coordinates for determining k, as it's a direct transformation of the data.
  # First, perform a preliminary PCA to get coordinates for k-determination.
  # Scale = TRUE is important for PCA if features are on different scales, but here they are binary (0/1).
  # Center = TRUE is usually good practice.
  # Ensure there are enough features (columns) for PCA to produce at least 2 components
  if (ncol(presence_absence_matrix) < 2) {
      message("Not enough unique terms (columns) for PCA to yield 2 dimensions. Cannot determine optimal k via silhouette on PCA. Defaulting to k=1.")
      optimal_k <- 1
  } else {
      pca_prelim <- prcomp(presence_absence_matrix, center = TRUE, scale. = FALSE)
      # Use the first two principal components for k-determination
      pca_coords_for_k <- as.data.frame(pca_prelim$x[, 1:2])

      if (nrow(pca_coords_for_k) < 2) {
          message("Not enough unique data points after PCA for robust K-means or silhouette analysis. Defaulting to k=1.")
          optimal_k <- 1
      } else {
          # Compute and plot optimal k using silhouette method
          fviz_nbclust_plot <- fviz_nbclust(pca_coords_for_k, kmeans, method = "silhouette", k.max = k_max)
          pdf("Kmeans_Optimal_K_Silhouette.pdf", width = 8, height = 6)
          print(fviz_nbclust_plot)
          dev.off()
          message("Optimal K-means clusters plot saved as 'Kmeans_Optimal_K_Silhouette.pdf'")

          # You can manually adjust optimal_k based on 'Kmeans_Optimal_K_Silhouette.pdf'
          # For example: optimal_k <- 4
          # The current default (3) is a placeholder.
      }
  }
}

message(paste("Using k = ", optimal_k, " for K-means clustering.", sep = ""))


# --- Step 5: Dimensionality Reduction using Multi-dimensional Scaling (MDS) (using Euclidean Distance) ---

message("Performing Multi-dimensional Scaling (MDS) (using Euclidean Distance)...")

# Perform Classical MDS (Principal Coordinates Analysis) using Euclidean distance
mds_result <- cmdscale(euclidean_distance_matrix, k = 2, eig = TRUE) # Changed input distance matrix

# Extract the coordinates for plotting
mds_coordinates <- as.data.frame(mds_result$points)
colnames(mds_coordinates) <- c("Dim1", "Dim2")
mds_coordinates$Database <- rownames(mds_coordinates)

# Perform K-means clustering on the MDS coordinates
if (optimal_k > 0 && nrow(mds_coordinates) >= optimal_k) { # Ensure enough points for k-means
  kmeans_mds_result <- kmeans(mds_coordinates[, c("Dim1", "Dim2")], centers = optimal_k, nstart = 25)
  mds_coordinates$Cluster <- as.factor(kmeans_mds_result$cluster)
} else {
  mds_coordinates$Cluster <- as.factor(rep(1, nrow(mds_coordinates))) # Assign all to one cluster if k-means not feasible
}


# --- Step 6: Dimensionality Reduction using Principal Component Analysis (PCA) ---

message("Performing Principal Component Analysis (PCA)...")

# Perform PCA on the presence-absence matrix
pca_result <- prcomp(presence_absence_matrix, center = TRUE, scale. = FALSE)

# Extract the principal components for plotting (typically PC1 and PC2)
pca_coordinates <- as.data.frame(pca_result$x[, 1:2]) # Get first two principal components
colnames(pca_coordinates) <- c("PC1", "PC2")
pca_coordinates$Database <- rownames(pca_coordinates)

# Add K-means cluster assignments to PCA coordinates for consistent plotting
if (optimal_k > 0 && nrow(pca_coordinates) >= optimal_k) { # Ensure enough points for k-means
  kmeans_pca_result <- kmeans(pca_coordinates[, c("PC1", "PC2")], centers = optimal_k, nstart = 25)
  pca_coordinates$Cluster <- as.factor(kmeans_pca_result$cluster)
} else {
  pca_coordinates$Cluster <- as.factor(rep(1, nrow(pca_coordinates))) # Assign all to one cluster if k-means not feasible
}


# --- Step 7: Visualize the relationships in 2D space (MDS and PCA Plots) ---

message("Generating MDS and PCA plots...")

# MDS Plot (using Euclidean Distance)
mds_plot <- ggplot(mds_coordinates, aes(x = Dim1, y = Dim2, color = Cluster, label = Database)) +
  geom_point(size = 4, alpha = 0.8) +
  geom_text(vjust = -1, hjust = 0.5, size = 3) +
  theme_minimal() +
  labs(
    title = "2D MDS Projection of Databases (Euclidean Distance)", # Changed title
    subtitle = paste("Clustered using K-means (k=", optimal_k, ")", sep = ""),
    x = "MDS Dimension 1",
    y = "MDS Dimension 2"
  ) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold"),
    plot.subtitle = element_text(hjust = 0.5),
    legend.title = element_text(face = "bold")
  )

pdf("MDS_Projection_Clustered_Euclidean.pdf", width = 10, height = 8) # Changed filename
print(mds_plot)
dev.off()
message("MDS projection plot saved as 'MDS_Projection_Clustered_Euclidean.pdf'")


# PCA Plot
pca_plot <- ggplot(pca_coordinates, aes(x = PC1, y = PC2, color = Cluster, label = Database)) +
  geom_point(size = 4, alpha = 0.8) +
  geom_text(vjust = -1, hjust = 0.5, size = 3) +
  theme_minimal() +
  labs(
    title = "2D PCA Projection of Databases (Presence-Absence Data)",
    subtitle = paste("Clustered using K-means (k=", optimal_k, ")", sep = ""),
    x = paste0("Principal Component 1 (", round(summary(pca_result)$importance[2,1]*100, 2), "%)"),
    y = paste0("Principal Component 2 (", round(summary(pca_result)$importance[2,2]*100, 2), "%)")
  ) +
  theme(
    plot.title = element_text(hjust = 0.5, face = "bold"),
    plot.subtitle = element_text(hjust = 0.5),
    legend.title = element_text(face = "bold")
  )

pdf("PCA_Projection_Clustered.pdf", width = 10, height = 8)
print(pca_plot)
dev.off()
message("PCA projection plot saved as 'PCA_Projection_Clustered.pdf'")


# --- NEW: Generate Unique Terms Summary Table ---
message("\n--- Generating Unique Terms Summary Table ---")

# Initialize an empty data frame for the summary table
summary_table_data <- data.frame(
  "Sheet Name" = character(),
  "Herbs Count" = integer(),
  "Formulas Count" = integer(),
  "Targets Count" = integer(),
  "Diseases Count" = integer(),
  "Ingredients Count" = integer(),
  "Total Unique Terms (Sheet)" = integer(),
  stringsAsFactors = FALSE
)

# Populate the summary table
if (length(list_of_term_counts_per_column_per_sheet) > 0) {
  for (sheet_name in names(list_of_term_counts_per_column_per_sheet)) {
    col_counts <- list_of_term_counts_per_column_per_sheet[[sheet_name]]
    total_unique_for_sheet <- length(list_of_term_sets[[sheet_name]]) # Total unique terms for the sheet (across all 5 columns)

    row_data_list <- list(
      "Sheet Name" = sheet_name,
      "Herbs Count" = col_counts[["Herbs"]],
      "Formulas Count" = col_counts[["Formulas"]],
      "Targets Count" = col_counts[["Targets"]],
      "Diseases Count" = col_counts[["Diseases"]],
      "Ingredients Count" = col_counts[["Ingredients"]],
      "Total Unique Terms (Sheet)" = total_unique_for_sheet
    )
    summary_table_data <- rbind(summary_table_data, as.data.frame(row_data_list, stringsAsFactors = FALSE))
  }

  # Print the table to PDF
  pdf("Unique_Terms_Summary_Table.pdf", width = 12, height = 8) # Adjust width/height as needed
  # Use tableGrob from gridExtra for plotting data frames as tables
  # Adjust theme for better appearance if needed, e.g., theme_minimal()
  grid.table(summary_table_data, rows = NULL, theme = ttheme_default(base_size = 8)) # base_size for text size
  dev.off()
  message("Unique terms summary table saved as 'Unique_Terms_Summary_Table.pdf'")
} else {
  message("No terms found across any sheets to generate a unique terms summary table.")
}


message("Analysis complete. Check the generated PDF files for results:")
message(" - Hierarchical_Clustering_Dendrogram_Euclidean.pdf (New dendrogram using Euclidean distance)")
message(" - Kmeans_Optimal_K_Silhouette.pdf (Optimal K for K-means)")
message(" - MDS_Projection_Clustered_Euclidean.pdf (New MDS plot using Euclidean distance)")
message(" - PCA_Projection_Clustered.pdf (PCA plot)")
message(" - Jaccard_Similarity_Matrix_Heatmap.pdf (Jaccard Similarity Matrix)")
message(" - Unique_Terms_Summary_Table.pdf (Unique terms count summary)")

OUTPUT MATRIX FINAL

# --- Install necessary packages if you haven't already ---
# Uncomment and run these lines if you get errors about missing packages.
# It's crucial that these installations complete successfully before running the rest of the script.
# install.packages("readxl")
# install.packages("dplyr")
# install.packages("tidyr")
# install.packages("corrplot")
# install.packages("gridExtra")
# install.packages("gtable")
# install.packages("grid")
# install.packages("cluster") # For divisive hierarchical clustering (diana function) and silhouette
# install.packages("factoextra") # For enhanced dendrogram plotting, and some clustering utilities
# install.packages("writexl") # This is needed if you want to modify Excel files programmatically
# install.packages("klaR") # For K-modes clustering
# install.packages("ggplot2") # For plotting the optimal K results

# Load the required libraries
library(readxl)
library(dplyr)
library(tidyr)
library(corrplot)
library(gridExtra)
library(gtable)
library(grid)
library(cluster) # Load the cluster package
library(factoextra) # Load the factoextra package
library(klaR) # Load the klaR package for kmodes
library(ggplot2) # Load ggplot2 for plotting optimal K

# --- Function to calculate and plot Jaccard Similarity and Overlap Counts Matrices ---
calculate_and_plot_matrices <- function(combined_data_list, output_dir = "./Output_Matrices/") {

  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
  }

  for (col_name in names(combined_data_list)) {
    df_binary <- combined_data_list[[col_name]]

    if (is.null(df_binary) || nrow(df_binary) == 0 || ncol(df_binary) < 2) {
      message(paste0("Skipping matrix analysis for category '", col_name, "': No data or insufficient data source columns (need at least 1 besides 'Value')."))
      next
    }
    
    if (!"Value" %in% colnames(df_binary)) {
      message(paste0("WARNING: Sheet '", col_name, "' is missing the 'Value' column. Cannot proceed with matrix analysis."))
      next
    }

    source_cols <- setdiff(colnames(df_binary), "Value")

    if (length(source_cols) < 2) {
        message(paste0("Skipping matrix analysis for category '", col_name, "': Only ", length(source_cols), " data source columns found after processing. Need at least 2 for comparisons."))
        next
    }

    binary_matrix <- as.matrix(df_binary[, source_cols])

    if (all(binary_matrix == 0)) {
      message(paste0("Skipping matrix analysis for category '", col_name, "': All elements are 0 across all sheets. No overlaps or unique elements to calculate."))
      next
    }

    jaccard_dist_matrix <- dist(t(binary_matrix), method = "binary")
    jaccard_sim_matrix <- 1 - as.matrix(jaccard_dist_matrix)
    diag(jaccard_sim_matrix) <- 1
    jaccard_sim_matrix[is.nan(jaccard_sim_matrix)] <- 0

    csv_file_name_jaccard <- file.path(output_dir, paste0("Jaccard_Similarity_Matrix_", gsub(" ", "_", col_name), ".csv"))
    write.csv(jaccard_sim_matrix, csv_file_name_jaccard, row.names = TRUE)
    message(paste0("Saved Jaccard similarity matrix for '", col_name, "' to '", csv_file_name_jaccard, "'"))

    pdf_file_name_jaccard <- file.path(output_dir, paste0("Jaccard_Similarity_Matrix_", gsub(" ", "_", col_name), ".pdf"))
    
    message(paste0("  DEBUG: Plotting '", col_name, "' Jaccard Similarity Matrix. Dimensions: ",
                   paste(dim(jaccard_sim_matrix), collapse = "x"),
                   ", Range: [", round(min(jaccard_sim_matrix, na.rm = TRUE), 4), ", ", round(max(jaccard_sim_matrix, na.rm = TRUE), 4), "]",
                   ", All zeros: ", all(jaccard_sim_matrix == 0)))

    tryCatch({
      pdf(pdf_file_name_jaccard, width = 10, height = 10)
      
      if (nrow(jaccard_sim_matrix) >= 2 && ncol(jaccard_sim_matrix) >= 2) {
        corrplot(jaccard_sim_matrix,
                 method = "color",
                 type = "full",
                 order = "hclust",
                 tl.col = "black",
                 tl.srt = 45,
                 addCoef.col = "black",
                 number.cex = 0.7,
                 main = paste("Jaccard Similarity Matrix for", col_name),
                 col = colorRampPalette(c("white", "lightblue", "darkblue"))(100)
        )
      } else {
        grid.newpage()
        grid.text(paste("Jaccard Similarity Matrix for", col_name, "\n(Not suitable for heatmap, showing as table)"),
                  y = 0.95, gp = gpar(fontsize = 12, fontface = "bold"))
        
        display_data <- round(jaccard_sim_matrix, 4)
        if (nrow(display_data) >= 1 && ncol(display_data) >= 1) {
            table_grob <- tableGrob(
                display_data,
                rows = rownames(display_data),
                cols = colnames(display_data),
                theme = ttheme_default(base_size = 8)
            )
            grid.draw(table_grob)
        } else {
            grid.text("No valid data for table display.", y = 0.5, gp = gpar(fontsize = 10))
        }
      }
      message(paste0("Successfully generated plot for '", col_name, "' to '", pdf_file_name_jaccard, "'"))
    }, error = function(e) {
      message(paste0("ERROR: Failed to generate PDF for '", col_name, "' (Jaccard Similarity). Error: ", e$message))
    }, finally = {
      if (dev.cur() != 1) {
        dev.off()
      }
    })

    overlap_counts_matrix <- t(binary_matrix) %*% binary_matrix

    csv_file_name_overlap_counts <- file.path(output_dir, paste0("Overlap_Counts_Matrix_", gsub(" ", "_", col_name), ".csv"))
    write.csv(overlap_counts_matrix, csv_file_name_overlap_counts, row.names = TRUE)
    message(paste0("Saved Overlap Counts matrix for '", col_name, "' to '", csv_file_name_overlap_counts, "'"))

    pdf_file_name_intersection_table <- file.path(output_dir, paste0("Intersection_Size_Table_", gsub(" ", "_", col_name), ".pdf"))
    
    message(paste0("  DEBUG: Plotting '", col_name, "' Intersection Counts Table. Dimensions: ",
                   paste(dim(overlap_counts_matrix), collapse = "x"),
                   ", Range: [", round(min(overlap_counts_matrix, na.rm = TRUE), 4), ", ", round(max(overlap_counts_matrix, na.rm = TRUE), 4), "]",
                   ", All zeros: ", all(overlap_counts_matrix == 0)))

    tryCatch({
      pdf(pdf_file_name_intersection_table, width = 10, height = 10)
      
      grid.newpage()
      grid.text(paste("Pairwise Intersection Counts for", col_name),
                y = 0.95, gp = gpar(fontsize = 12, fontface = "bold"))
      
      display_data <- overlap_counts_matrix
      if (nrow(display_data) >= 1 && ncol(display_data) >= 1) {
            table_grob <- tableGrob(
                display_data,
                rows = rownames(display_data),
                cols = colnames(display_data),
                theme = ttheme_default(base_size = 8)
            )
            grid.draw(table_grob)
        } else {
            grid.text("No Intersection Counts data available for table display.", y = 0.5, gp = gpar(fontsize = 12))
        }
      message(paste0("Successfully generated plot for '", col_name, "' to '", pdf_file_name_intersection_table, "'"))
    }, error = function(e) {
      message(paste0("ERROR: Failed to generate PDF for '", col_name, "' (Intersection Counts Table). Error: ", e$message))
    }, finally = {
      if (dev.cur() != 1) {
        dev.off()
      }
    })

    pdf_file_name_overlap <- file.path(output_dir, paste0("Overlap_Counts_Matrix_Scaled_", gsub(" ", "_", col_name), ".pdf"))
    
    max_val_overlap <- max(overlap_counts_matrix, na.rm = TRUE)
    if (is.infinite(max_val_overlap) || is.na(max_val_overlap) || max_val_overlap == 0) {
        scaled_overlap_counts_matrix <- matrix(0, nrow = nrow(overlap_counts_matrix), ncol = ncol(overlap_counts_matrix),
                                               dimnames = dimnames(overlap_counts_matrix))
    } else {
        scaled_overlap_counts_matrix <- overlap_counts_matrix / max_val_overlap
    }

    message(paste0("  DEBUG: Plotting '", col_name, "' Overlap Counts Matrix Scaled. Dimensions: ",
                   paste(dim(scaled_overlap_counts_matrix), collapse = "x"),
                   ", Range: [", round(min(scaled_overlap_counts_matrix, na.rm = TRUE), 4), ", ", round(max(scaled_overlap_counts_matrix, na.rm = TRUE), 4), "]",
                   ", All zeros: ", all(scaled_overlap_counts_matrix == 0)))

    tryCatch({
      pdf(pdf_file_name_overlap, width = 10, height = 10)
      
      if (nrow(overlap_counts_matrix) >= 2 && ncol(overlap_counts_matrix) >= 2) {
        if (all(scaled_overlap_counts_matrix == 0)) {
            grid.newpage()
            grid.text(paste("Pairwise Overlap Counts (Scaled 0-1) Matrix for", col_name, "\n(All overlaps are zero)"),
                      y = 0.5, gp = gpar(fontsize = 12, fontface = "bold"))
        } else {
            corrplot(scaled_overlap_counts_matrix,
                     method = "square",
                     type = "full",
                     order = "hclust",
                     tl.col = "black",
                     tl.srt = 45,
                     addCoef.col = "black",
                     number.cex = 0.7,
                     main = paste("Pairwise Overlap Counts (Scaled 0-1) Matrix for", col_name),
                     col = colorRampPalette(c("white", "lightcoral", "darkred"))(100),
                     cl.lim = c(0, 1),
                     cl.pos = "r"
            )
        }
      } else {
        grid.newpage()
        grid.text(paste("Pairwise Overlap Counts Matrix for", col_name, "\n(Not suitable for heatmap, showing as table)"),
                  y = 0.95, gp = gpar(fontsize = 12, fontface = "bold"))
        
        display_data <- overlap_counts_matrix
        if (nrow(display_data) >= 1 && ncol(display_data) >= 1) {
            table_grob <- tableGrob(
                display_data,
                rows = rownames(display_data),
                cols = colnames(display_data),
                theme = ttheme_default(base_size = 8)
            )
            grid.draw(table_grob)
        } else {
            grid.text("No valid data for table display.", y = 0.5, gp = gpar(fontsize = 10))
        }
      }
      message(paste0("Successfully generated plot for '", col_name, "' to '", pdf_file_name_overlap, "'"))
    }, error = function(e) {
      message(paste0("ERROR: Failed to generate PDF for '", col_name, "' (Overlap Counts Scaled). Error: ", e$message))
    }, finally = {
      if (dev.cur() != 1) {
        dev.off()
      }
    })

    num_sheets <- nrow(overlap_counts_matrix)
    overlap_percentage_matrix <- matrix(0, nrow = num_sheets, ncol = num_sheets,
                                           dimnames = dimnames(overlap_counts_matrix))

    for (i in 1:num_sheets) {
      for (j in 1:num_sheets) {
        denominator_size <- overlap_counts_matrix[j,j]
        if (denominator_size > 0) {
          overlap_percentage_matrix[i,j] <- (overlap_counts_matrix[i,j] / denominator_size) * 100
        } else {
          overlap_percentage_matrix[i,j] <- 0
        }
      }
    }
    overlap_percentage_matrix[is.nan(overlap_percentage_matrix)] <- 0

    csv_file_name_overlap_percent <- file.path(output_dir, paste0("Overlap_Percentage_Matrix_", gsub(" ", "_", col_name), ".csv"))
    write.csv(overlap_percentage_matrix, csv_file_name_overlap_percent, row.names = TRUE)
    message(paste0("Saved Overlap Percentage matrix for '", col_name, "' to '", csv_file_name_overlap_percent, "'"))

    pdf_file_name_overlap_percent <- file.path(output_dir, paste0("Overlap_Percentage_Matrix_", gsub(" ", "_", col_name), ".pdf"))
    
    scaled_overlap_percentage_matrix_for_plot <- overlap_percentage_matrix / 100

    message(paste0("  DEBUG: Plotting '", col_name, "' Overlap Percentage Matrix Scaled. Dimensions: ",
                   paste(dim(scaled_overlap_percentage_matrix_for_plot), collapse = "x"),
                   ", Range: [", round(min(scaled_overlap_percentage_matrix_for_plot, na.rm = TRUE), 4), ", ", round(max(scaled_overlap_percentage_matrix_for_plot, na.rm = TRUE), 4), "]",
                   ", All zeros: ", all(scaled_overlap_percentage_matrix_for_plot == 0)))

    tryCatch({
      pdf(pdf_file_name_overlap_percent, width = 10, height = 10)

      if (nrow(overlap_percentage_matrix) >= 2 && ncol(overlap_percentage_matrix) >= 2) {
        corrplot(scaled_overlap_percentage_matrix_for_plot,
                 method = "color",
                 type = "full",
                 order = "hclust",
                 tl.col = "black",
                 tl.srt = 45,
                 addCoef.col = "black",
                 number.cex = 0.7,
                 main = paste("Overlap Percentage (Intersection / Column Set Size) Matrix for", col_name, " (Scaled 0-1)"),
                 col = colorRampPalette(c("white", "lightgreen", "darkgreen"))(100),
                 cl.lim = c(0, 1),
                 cl.pos = "r"
        )
      } else {
        grid.newpage()
        grid.text(paste("Overlap Percentage Matrix for", col_name, "\n(Not suitable for heatmap, showing as table)"),
                  y = 0.95, gp = gpar(fontsize = 12, fontface = "bold"))
        
        display_data <- round(overlap_percentage_matrix, 2)
        if (nrow(display_data) >= 1 && ncol(display_data) >= 1) {
            table_grob <- tableGrob(
                display_data,
                rows = rownames(display_data),
                cols = colnames(display_data),
                theme = ttheme_default(base_size = 8)
            )
            grid.draw(table_grob)
        } else {
            grid.text("No valid data for table display.", y = 0.5, gp = gpar(fontsize = 10))
        }
      }
      message(paste0("Successfully generated plot for '", col_name, "' to '", pdf_file_name_overlap_percent, "'"))
    }, error = function(e) {
      message(paste0("ERROR: Failed to generate PDF for '", col_name, "' (Overlap Percentage). Error: ", e$message))
    }, finally = {
      if (dev.cur() != 1) {
        dev.off()
      }
    })

  }
}

# --- Function to perform and plot both Agglomerative and Divisive Clustering ---
perform_hierarchical_clustering <- function(jaccard_matrix_path, output_dir = "./Output_Matrices/") {
  
  # Extract category name from file path for titles
  file_name <- basename(jaccard_matrix_path)
  category_name <- gsub("Jaccard_Similarity_Matrix_|.csv", "", file_name)
  plot_title_suffix <- paste("Databases based on", gsub("_", " ", category_name), "Jaccard Similarity")

  # 1. Load the Jaccard Similarity Matrix
  if (!file.exists(jaccard_matrix_path)) {
    message(paste0("Skipping clustering for '", category_name, "': Jaccard matrix file not found at '", jaccard_matrix_path, "'"))
    return(NULL)
  }
  similarity_matrix <- read.csv(jaccard_matrix_path, row.names = 1)
  similarity_matrix <- as.matrix(similarity_matrix)

  # Check for valid dimensions for clustering (at least 2x2)
  if (nrow(similarity_matrix) < 2 || ncol(similarity_matrix) < 2) {
    message(paste0("Skipping clustering for '", category_name, "': Matrix is too small (", nrow(similarity_matrix), "x", ncol(similarity_matrix), "). Needs at least 2x2 for hierarchical clustering."))
    return(NULL)
  }
  
  # Ensure all values are numeric and handle any potential non-numeric conversion issues
  if (!all(sapply(similarity_matrix, is.numeric))) {
    message(paste0("Skipping clustering for '", category_name, "': Matrix contains non-numeric values after loading. Please check the CSV file."))
    return(NULL)
  }

  # 2. Convert Jaccard Similarity to Distance
  distance_matrix <- as.dist(1 - similarity_matrix)

  # --- Agglomerative Hierarchical Clustering (Bottom-Up) ---
  message(paste0("\n--- Performing Agglomerative Clustering for ", category_name, " ---"))
  hc_agglomerative <- hclust(distance_matrix, method = "ward.D2") # 'ward.D2' is a common and often effective method

  pdf_agglomerative_file <- file.path(output_dir, paste0("Dendrogram_Agglomerative_", gsub(" ", "_", category_name), ".pdf"))
  tryCatch({
    pdf(pdf_agglomerative_file, width = 12, height = 8)
    plot(hc_agglomerative, 
         main = paste("Agglomerative Clustering of", plot_title_suffix),
         xlab = "Databases", 
         ylab = "Distance (1 - Jaccard Similarity)",
         hang = -1 # Align labels at the bottom
    )
    # Optional: Draw rectangles around clusters (e.g., 3 clusters)
    # rect.hclust(hc_agglomerative, k = 3, border = 2:4)
    message(paste0("Saved Agglomerative Dendrogram for '", category_name, "' to '", pdf_agglomerative_file, "'"))
  }, error = function(e) {
    message(paste0("ERROR: Failed to plot Agglomerative Dendrogram for '", category_name, "'. Error: ", e$message))
  }, finally = {
    if (dev.cur() != 1) { dev.off() }
  })
  
  # --- Divisive Hierarchical Clustering (Top-Down) ---
  message(paste0("\n--- Performing Divisive Clustering for ", category_name, " ---"))
  # 'diana' function from the 'cluster' package performs divisive clustering
  hc_divisive <- diana(distance_matrix) 

  pdf_divisive_file <- file.path(output_dir, paste0("Dendrogram_Divisive_", gsub(" ", "_", category_name), ".pdf"))
  tryCatch({
    pdf(pdf_divisive_file, width = 12, height = 8)
    # Plotting diana results is similar to hclust
    plot(hc_divisive, 
         main = paste("Divisive Clustering of", plot_title_suffix),
         xlab = "Databases", 
         ylab = "Distance (1 - Jaccard Similarity)",
         hang = -1
    )
    # Optional: Draw rectangles around clusters (e.g., 3 clusters)
    # rect.hclust(as.hclust(hc_divisive), k = 3, border = 2:4) # diana object needs conversion for rect.hclust
    message(paste0("Saved Divisive Dendrogram for '", category_name, "' to '", pdf_divisive_file, "'"))
  }, error = function(e) {
    message(paste0("ERROR: Failed to plot Divisive Dendrogram for '", category_name, "'. Error: ", e$message))
  }, finally = {
    if (dev.cur() != 1) { dev.off() }
  })

  # --- Extracting Cluster Assignments (Optional, but useful for analysis) ---
  num_clusters <- 3 # Default number of clusters for output
  if (nrow(similarity_matrix) >= num_clusters) { # Only cut if enough data points for specified clusters
    agglomerative_clusters <- cutree(hc_agglomerative, k = num_clusters)
    divisive_clusters <- cutree(as.hclust(hc_divisive), k = num_clusters) # diana object needs conversion for cutree
    
    message(paste0("\n--- Cluster Assignments (k=", num_clusters, ") for ", category_name, " ---"))
    message("Agglomerative Clusters:")
    print(agglomerative_clusters)
    message("\nDivisive Clusters:")
    print(divisive_clusters)
    
    write.csv(data.frame(Database = names(agglomerative_clusters), Agglomerative_Cluster = agglomerative_clusters), 
              file.path(output_dir, paste0("Clusters_Agglomerative_", gsub(" ", "_", category_name), "_k", num_clusters, ".csv")), row.names = FALSE)
    write.csv(data.frame(Database = names(divisive_clusters), Divisive_Cluster = divisive_clusters), 
              file.path(output_dir, paste0("Clusters_Divisive_", gsub(" ", "_", category_name), "_k", num_clusters, ".csv")), row.names = FALSE)
    message(paste0("Saved cluster assignments for '", category_name, "' to CSVs."))
  } else {
    message(paste0("Not enough databases to form ", num_clusters, " clusters for '", category_name, "'. Skipping cluster assignment printout."))
  }
}

# --- NEW: Function to find optimal K for K-modes using Silhouette method ---
find_optimal_k_kmodes <- function(kmodes_data_frame, jaccard_dist, category_name, output_dir, max_k_to_test = 10) {
  message(paste0("\n--- Finding Optimal K for K-modes (Silhouette Method) for ", category_name, " ---"))
  
  num_databases <- nrow(kmodes_data_frame)
  
  # Determine sensible range for k
  if (num_databases < 2) {
    message("  Not enough data points to find optimal K (need at least 2).")
    return(NULL)
  }
  
  # Max k should be less than number of data points
  k_range <- 2:min(max_k_to_test, num_databases - 1) 
  
  if (length(k_range) < 1) {
    message(paste0("  Cannot test a range of K values. Only ", num_databases, " databases available. Consider using k=2 directly."))
    return(NULL)
  }

  avg_silhouette_scores <- c()
  
  for (k in k_range) {
    message(paste0("    Testing k = ", k, " for ", category_name))
    kmodes_result <- tryCatch(
      kmodes(kmodes_data_frame, modes = k, iter.max = 100),
      error = function(e) {
        message(paste0("      Error running kmodes for k = ", k, ": ", e$message))
        return(NULL)
      }
    )
    
    if (!is.null(kmodes_result)) {
      # Ensure cluster assignments are named correctly for silhouette
      cluster_assignments <- kmodes_result$cluster
      names(cluster_assignments) <- rownames(kmodes_data_frame)
      
      # Calculate silhouette score. The silhouette function needs the distance matrix
      # and the cluster assignments. It's crucial that the order of data points
      # in the distance matrix matches the order in cluster_assignments.
      # The 'jaccard_dist' already reflects distances between databases.
      sil <- silhouette(cluster_assignments, dist = jaccard_dist)
      avg_silhouette_scores <- c(avg_silhouette_scores, mean(sil[, "sil_width"]))
    } else {
      avg_silhouette_scores <- c(avg_silhouette_scores, NA) # Mark as NA if clustering failed
    }
  }
  
  results_df <- data.frame(k = k_range, avg_silhouette_width = avg_silhouette_scores)
  
  # Plotting the silhouette scores
  plot_optimal_k_file <- file.path(output_dir, paste0("Optimal_K_Silhouette_KModes_", gsub(" ", "_", category_name), ".pdf"))
  tryCatch({
    pdf(plot_optimal_k_file, width = 8, height = 6)
    p <- ggplot(results_df, aes(x = k, y = avg_silhouette_width)) +
      geom_line(color = "steelblue") +
      geom_point(color = "steelblue") +
      geom_vline(xintercept = results_df$k[which.max(results_df$avg_silhouette_width)], 
                 linetype = "dashed", color = "red") +
      labs(title = paste("Optimal K for K-modes (Silhouette Method) for", category_name),
           subtitle = paste0("Suggested K: ", results_df$k[which.max(results_df$avg_silhouette_width)], " (Max Avg. Silhouette Width)"),
           x = "Number of Clusters (k)",
           y = "Average Silhouette Width") +
      theme_minimal()
    print(p)
    message(paste0("  Saved Optimal K plot for '", category_name, "' to '", plot_optimal_k_file, "'"))
  }, error = function(e) {
    message(paste0("  ERROR: Failed to plot Optimal K for '", category_name, "'. Error: ", e$message))
  }, finally = {
    if (dev.cur() != 1) {
      dev.off()
    }
  })
  
  # Determine optimal k based on highest average silhouette width
  optimal_k <- results_df$k[which.max(results_df$avg_silhouette_width)]
  message(paste0("  Optimal K for '", category_name, "' determined as: ", optimal_k, " (based on max average silhouette width)."))
  
  return(optimal_k)
}


# --- Function to perform MDS and K-modes clustering ---
perform_mds_and_kmodes <- function(df_binary, category_name, output_dir = "./Output_Matrices/") {
  message(paste0("\n--- Performing MDS and K-modes Clustering for ", category_name, " ---"))

  if (is.null(df_binary) || nrow(df_binary) == 0) {
    message(paste0("Skipping MDS and K-modes for '", category_name, "': No data available."))
    return(NULL)
  }

  # Ensure 'Value' column exists before trying to remove it
  if (!"Value" %in% colnames(df_binary)) {
    message(paste0("WARNING: Sheet '", category_name, "' is missing the 'Value' column. Cannot proceed with MDS/K-modes."))
    return(NULL)
  }
  
  # Identify source columns (all columns except 'Value')
  source_cols <- setdiff(colnames(df_binary), "Value")

  # Ensure there are at least 2 source columns for distance calculation and MDS
  if (length(source_cols) < 2) {
    message(paste0("Skipping MDS and K-modes for '", category_name, "': Not enough source columns (need at least 2) after cleaning. Found: ", length(source_cols)))
    return(NULL)
  }
  
  # Extract the binary matrix (transpose for column-wise comparison of databases for MDS)
  binary_matrix_for_dist <- as.matrix(df_binary[, source_cols])

  # Check if all values in the binary matrix are zero
  if (all(binary_matrix_for_dist == 0)) {
    message(paste0("Skipping MDS and K-modes for '", category_name, "': All binary values are 0. No variation for clustering/MDS."))
    return(NULL)
  }

  # 1. Calculate Jaccard Distance for MDS
  # Since we want to cluster the *databases* (columns), we transpose the matrix *before* calculating distance.
  # dist(t(binary_matrix_for_dist)) correctly calculates distances between columns of original df_binary.
  jaccard_dist <- dist(t(binary_matrix_for_dist), method = "binary")

  # Check for problematic distances (e.g., all 0 or infinite)
  if (any(is.na(jaccard_dist)) || all(jaccard_dist == 0)) {
      message(paste0("Skipping MDS for '", category_name, "': Jaccard distances are problematic (all identical or contain NA)."))
      return(NULL)
  }

  # 2. Perform MDS
  mds_result <- cmdscale(jaccard_dist, k = 2) # k=2 for 2D plot
  colnames(mds_result) <- c("Dim1", "Dim2")
  mds_df <- as.data.frame(mds_result)
  mds_df$Database <- rownames(mds_df)

  # Check if MDS produced valid results (not all NA or Inf)
  if (all(is.na(mds_result)) || all(is.infinite(mds_result))) {
      message(paste0("Skipping MDS plot for '", category_name, "': MDS resulted in all NA/Inf coordinates."))
      return(NULL)
  }
  
  # 3. Prepare data for K-modes clustering and Optimal K determination
  # K-modes should be performed on the **transposed binary matrix** so that
  # each ROW represents a database (the object to be clustered)
  # and each COLUMN represents a 'Value' (the features based on which databases are clustered).
  kmodes_input_matrix <- t(binary_matrix_for_dist)
  
  # K-modes requires a data frame with factors or characters
  kmodes_data_frame <- as.data.frame(kmodes_input_matrix)
  kmodes_data_frame[] <- lapply(kmodes_data_frame, factor, levels = c("0", "1"))

  # --- NEW: Find Optimal K using Silhouette method ---
  optimal_k <- find_optimal_k_kmodes(kmodes_data_frame, jaccard_dist, category_name, output_dir)
  
  if (is.null(optimal_k)) {
    message(paste0("K-modes optimal K determination failed or insufficient data for '", category_name, "'. Skipping K-modes clustering."))
    mds_df$Cluster <- "Not Clustered" # Placeholder for plotting
    k_modes_clusters_used <- "N/A" # For plot title
  } else {
    # If optimal_k is found, use it for clustering
    k_modes_clusters_used <- optimal_k
    num_databases <- nrow(kmodes_data_frame) # Re-get num_databases in case adjusted earlier
    
    if (num_databases < k_modes_clusters_used) {
        # This case should ideally be handled by find_optimal_k_kmodes, but as a safeguard
        message(paste0("Adjusted K-modes clusters to ", max(2, num_databases - 1), " for '", category_name, "' as optimal K (", k_modes_clusters_used, ") was too high for ", num_databases, " databases."))
        k_modes_clusters_used <- max(2, num_databases - 1)
    }
    
    kmodes_result <- tryCatch(
        kmodes(kmodes_data_frame, modes = k_modes_clusters_used, iter.max = 100),
        error = function(e) {
            message(paste0("Error in kmodes for '", category_name, "' (k=", k_modes_clusters_used, "): ", e$message))
            return(NULL)
        }
    )

    if (!is.null(kmodes_result)) {
      cluster_assignments <- kmodes_result$cluster
      names(cluster_assignments) <- rownames(kmodes_data_frame)
      mds_df$Cluster <- as.factor(cluster_assignments[mds_df$Database])
      
      csv_file_name_kmodes_clusters <- file.path(output_dir, paste0("KModes_Clusters_", gsub(" ", "_", category_name), "_k", k_modes_clusters_used, ".csv"))
      write.csv(mds_df[, c("Database", "Cluster")], csv_file_name_kmodes_clusters, row.names = FALSE)
      message(paste0("Saved K-modes cluster assignments for '", category_name, "' to '", csv_file_name_kmodes_clusters, "'"))
    } else {
      mds_df$Cluster <- "Not Clustered"
      message(paste0("K-modes clustering skipped/failed for '", category_name, "'. MDS plot will not show clusters."))
    }
  }

  # 4. Plot MDS results with K-modes clusters (using the determined K)
  pdf_mds_file <- file.path(output_dir, paste0("MDS_KModes_Plot_", gsub(" ", "_", category_name), "_k", k_modes_clusters_used, ".pdf"))
  
  message(paste0("  DEBUG: Plotting '", category_name, "' MDS with K-modes. Dimensions: ",
                 paste(dim(mds_df), collapse = "x")))

  tryCatch({
    pdf(pdf_mds_file, width = 10, height = 8)
    
    if (nrow(mds_df) >= 2 && ncol(mds_df) >= 2) {
        plot_obj <- fviz_cluster(list(data = mds_df[, c("Dim1", "Dim2")], cluster = mds_df$Cluster),
                                 geom = c("point", "text"), # Include both points and text
                                 ellipse.type = "convex", # draws ellipse around clusters
                                 repel = TRUE, # avoids text overlapping
                                 main = paste("MDS Plot with K-modes Clusters (k=", k_modes_clusters_used, ") for", category_name),
                                 xlab = "Dimension 1", ylab = "Dimension 2"
                                 ) + theme_minimal()
        print(plot_obj)

    } else {
        grid.newpage()
        grid.text(paste("MDS Plot for", category_name, "\n(Not suitable for scatter plot, insufficient data)"),
                  y = 0.5, gp = gpar(fontsize = 12, fontface = "bold"))
    }
    message(paste0("Saved MDS-Kmodes plot for '", category_name, "' to '", pdf_mds_file, "'"))
  }, error = function(e) {
    message(paste0("ERROR: Failed to plot MDS-Kmodes for '", category_name, "'. Error: ", e$message))
  }, finally = {
    if (dev.cur() != 1) {
      dev.off()
    }
  })
}


# --- Main execution script ---

# IMPORTANT: Set this path to your "Combined data copy.xlsx" file
# Replace "path/to/your/Combined data copy.xlsx" with the actual path on your computer.
combined_data_excel_path <- "path/to/your/Combined data copy.xlsx" # <<<--- REPLACE THIS WITH YOUR ACTUAL FILE PATH

message("\n--- Starting process to load 'Combined data copy.xlsx' ---")

if (!file.exists(combined_data_excel_path)) {
  stop(paste0("Error: 'Combined data copy.xlsx' not found at ", combined_data_excel_path, ". Please ensure the file exists and the path is correct."))
}

category_sheet_names <- excel_sheets(combined_data_excel_path)
combined_results <- list()
message(paste0("Loading ", length(category_sheet_names), " sheets from '", basename(combined_data_excel_path), "'..."))
for (sheet_name in category_sheet_names) {
  message(paste0("  Loading category sheet: ", sheet_name))
  tryCatch({
    df <- read_excel(combined_data_excel_path, sheet = sheet_name, col_names = TRUE)
    
    # --- Data Cleaning and Binarization for each sheet ---
    
    # 1. Check for 'Value' column
    if (!"Value" %in% colnames(df)) {
      message(paste0("  WARNING for sheet '", sheet_name, "': 'Value' column not found. Skipping processing for this sheet."))
      combined_results[[sheet_name]] <- NULL # Mark as NULL so it's skipped later
      next # Skip to next sheet
    }

    # 2. Filter out columns that are entirely NA
    initial_cols_count <- ncol(df)
    df <- df[, colSums(is.na(df)) < nrow(df)]
    message(paste0("  Sheet '", sheet_name, "': After removing all-NA columns, ", ncol(df), " columns remaining (from ", initial_cols_count, ")."))

    # 3. Filter out rows where 'Value' is NA or empty string
    initial_rows_count <- nrow(df)
    df <- df %>% filter(!is.na(Value) & Value != "")
    message(paste0("  Sheet '", sheet_name, "': After filtering NA/empty 'Value' rows, ", nrow(df), " rows remaining (from ", initial_rows_count, ")."))

    # 4. Identify data source columns (all columns except 'Value')
    data_source_cols <- setdiff(colnames(df), "Value")

    # 5. Check if there are any data source columns left
    if (length(data_source_cols) == 0) {
      message(paste0("  WARNING for sheet '", sheet_name, "': No data source columns found besides 'Value' after cleaning. Skipping this sheet for matrix analysis."))
      combined_results[[sheet_name]] <- NULL # Mark as NULL
      next
    }
    
    # 6. Ensure all data source columns are treated as binary (1s or 0s)
    df_binary <- df %>%
      mutate(across(all_of(data_source_cols), ~ case_when(
        !is.na(.) & . != 0 ~ 1L, # Non-NA and non-zero -> 1 (integer)
        TRUE ~ 0L # Otherwise (NA or zero) -> 0 (integer)
      ))) %>%
      distinct() # Remove duplicate rows after binarization

    # 7. Store the processed (binarized) data
    combined_results[[sheet_name]] <- df_binary
    message(paste0("  Sheet '", sheet_name, "': Processed data dimensions: ", nrow(df_binary), " rows, ", ncol(df_binary), " columns."))

  }, error = function(e) {
    message(paste0("  Error processing sheet '", sheet_name, "': ", e$message, ". Skipping this sheet."))
    combined_results[[sheet_name]] <- NULL # Mark as NULL to skip
  })
}

message("\n--- Summary of Categories Loaded and Processed from 'Combined data copy.xlsx' ---")
if (length(combined_results) > 0) {
  for (col_name in names(combined_results)) {
    if (!is.null(combined_results[[col_name]]) && nrow(combined_results[[col_name]]) > 0) {
      cat(paste0("Category '", col_name, "': Successfully processed with ", nrow(combined_results[[col_name]]), " rows and ", ncol(combined_results[[col_name]]), " columns.\n"))
    } else {
      cat(paste0("Category '", col_name, "': No valid data after processing or failed to load.\n"))
    }
  }
} else {
  message("No categories or data found or loaded from 'Combined data copy.xlsx'.")
}

message("\n--- Calculating and Plotting Jaccard Similarity and Overlap Counts Matrices ---")
# Only pass sheets that were successfully processed
valid_results_for_matrices <- combined_results[!sapply(combined_results, is.null) & sapply(combined_results, nrow) > 0]
calculate_and_plot_matrices(valid_results_for_matrices, output_dir = "./Output_Matrices/")


message("\n--- Starting Hierarchical Clustering (Agglomerative & Divisive) Analysis ---")

# Define the output directory where your Jaccard matrices are saved
output_dir <- "./Output_Matrices/"

# Get a list of all generated Jaccard Similarity matrix files
jaccard_files <- list.files(output_dir, pattern = "^Jaccard_Similarity_Matrix_.*\\.csv$", full.names = TRUE)

if (length(jaccard_files) == 0) {
  message("No Jaccard Similarity Matrix CSV files found in the output directory. Please ensure the 'calculate_and_plot_matrices' function ran successfully.")
} else {
  for (file_path in jaccard_files) {
    perform_hierarchical_clustering(file_path, output_dir = output_dir)
  }
}

# --- Perform MDS and K-modes clustering for each processed sheet, now with optimal K determination ---
message("\n--- Starting MDS and K-modes Clustering Analysis (with Optimal K Determination) ---")

if (length(valid_results_for_matrices) == 0) {
    message("No valid sheets found for MDS and K-modes analysis.")
} else {
    for (sheet_name in names(valid_results_for_matrices)) {
        perform_mds_and_kmodes(
            df_binary = valid_results_for_matrices[[sheet_name]],
            category_name = sheet_name,
            output_dir = output_dir
            # Removed k_modes_clusters parameter here, as it's now determined automatically
        )
    }
}

message("\n--- All Analyses Complete. Check the 'Output_Matrices' folder for generated plots and CSVs. ---")
