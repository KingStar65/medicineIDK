#Combined data.xlsx

# Load necessary libraries
# If you don't have these installed, run:
# install.packages("readxl")   # For read_excel
# install.packages("dplyr")    # For data manipulation
# install.packages("tidyr")    # Explicitly for pivot_wider
# install.packages("writexl")  # For saving output to Excel

library(readxl)
library(dplyr)
library(tidyr) # Explicitly load tidyr for pivot_wider
library(writexl)

# --- Function to create combined unique data with sheet presence ---
create_combined_data_excel_r <- function(excel_file_path, sheet_names_list, output_file_name = "Combined data.xlsx") {
  
  if (!file.exists(excel_file_path)) {
    message(paste0("Error: Excel file not found at ", excel_file_path))
    return(list(tables = list())) # Return empty list if file not found
  }

  # Define the 5 specific column headers you are interested in
  target_columns <- c("Herbs", "Ingredients", "Formulas", "Targets", "Diseases")

  # Use a list of lists to collect data frames for each column type
  column_data_collectors <- list()
  for(col in target_columns) {
    column_data_collectors[[col]] <- list()
  }

  message(paste0("Processing ", length(sheet_names_list), " sheets from '", basename(excel_file_path), "'..."))

  for (sheet_name in sheet_names_list) {
    message(paste0("  Loading sheet: ", sheet_name))

    tryCatch({
      # Read the specific sheet from the Excel file
      df <- read_excel(excel_file_path, sheet = sheet_name, col_names = TRUE)

      # Process each target column
      for (col in target_columns) {
        if (col %in% names(df)) { # Check if column exists in the current sheet
          # Convert to character, lowercase, trim whitespace, and handle "NA"/"nan" strings
          processed_col_values <- tolower(trimws(as.character(df[[col]])))
          
          # Treat empty strings or "NA"/"nan" strings as actual NA, then remove NA values
          unique_vals <- unique(na.omit(processed_col_values[
            processed_col_values != "" &
            !is.na(processed_col_values) &
            processed_col_values != "na" & # Explicitly remove "na" string if it exists
            processed_col_values != "nan" # Explicitly remove "nan" string if it exists
          ]))

          # Create a temporary data frame for this column and sheet
          if (length(unique_vals) > 0) {
            temp_df <- data.frame(Value = unique_vals, Sheet = sheet_name, stringsAsFactors = FALSE)
          } else {
            temp_df <- data.frame(Value = character(0), Sheet = character(0), stringsAsFactors = FALSE)
          }

          # Add to the collector list for this column type
          column_data_collectors[[col]] <- c(column_data_collectors[[col]], list(temp_df))
        } else {
          message(paste0("    Warning: Column '", col, "' not found in sheet '", sheet_name, "'."))
        }
      }

    }, error = function(e) {
      message(paste0("Error loading or processing sheet '", sheet_name, "': ", e$message))
    })
  }
  
  # --- Combine all collected data for each target column ---
  # And prepare them for the detailed Excel output
  combined_data_tables <- list()

  for (col_name in target_columns) {
    # Combine all collected data frames for the current column
    if (length(column_data_collectors[[col_name]]) > 0) {
      all_data_for_col <- bind_rows(column_data_collectors[[col_name]])
    } else {
      all_data_for_col <- data.frame(Value = character(0), Sheet = character(0), stringsAsFactors = FALSE)
    }

    # Proceed only if the combined data frame for this column has any rows
    if (nrow(all_data_for_col) > 0) {
      # Get distinct Value-Sheet pairs (this is the core of deduplication and tracking presence)
      distinct_value_sheet_pairs <- all_data_for_col %>%
        distinct(Value, Sheet)

      # Pivot to wide format: Value as rows, Sheets as columns
      # *** CHANGED: Presence is now 1, and NA is filled with 0 ***
      pivoted_df <- distinct_value_sheet_pairs %>%
        mutate(Presence = 1L) %>% # Use 1L for integer 1
        pivot_wider(names_from = Sheet, values_from = Presence, values_fill = 0L) %>% # Use 0L for integer 0
        arrange(Value) # Sort by Value for clean output

      # Ensure 'Value' is the first column and all sheet columns are present and ordered
      # Create a template for all sheet columns, filling with 0L if a sheet has no data
      all_sheet_cols_template <- data.frame(matrix(0L, nrow = nrow(pivoted_df), ncol = length(sheet_names_list))) 
      colnames(all_sheet_cols_template) <- sheet_names_list
      
      # Fill the template with actual pivoted data (ensuring integer type)
      for (sn in sheet_names_list) {
        if (sn %in% names(pivoted_df)) {
          all_sheet_cols_template[[sn]] <- as.integer(pivoted_df[[sn]])
        }
      }
      
      # Combine Value column with the prepared numeric sheet columns
      final_pivoted_df <- bind_cols(select(pivoted_df, Value), all_sheet_cols_template)
      
      combined_data_tables[[col_name]] <- final_pivoted_df
    } else {
      # Create an empty data frame with appropriate columns if no data for this column
      empty_df <- data.frame(Value = character(0), stringsAsFactors = FALSE)
      for (sn in sheet_names_list) {
          empty_df[[sn]] <- integer(0) # Set to integer(0) for numeric type
      }
      combined_data_tables[[col_name]] <- empty_df
    }
  }
  
  # --- Save the results to the specified Excel file ---
  if (length(combined_data_tables) > 0 && any(sapply(combined_data_tables, function(df) nrow(df) > 0))) {
    writexl::write_xlsx(combined_data_tables, output_file_name)
    message(paste0("\nCombined unique data (with sheet occurrences) saved to '", output_file_name, "'"))
  } else {
    message(paste0("\nNo data to save to '", output_file_name, "'. This might mean no data was found or processing failed for all columns."))
  }

  return(combined_data_tables) # Return the list of dataframes for potential further use
}

# --- How to run the analysis on your machine ---

# IMPORTANT: Set this path to your "Compiled_done.xlsx" file
# Example for Windows: excel_file_path <- "C:/Users/YourUsername/Desktop/Compiled_done.xlsx"
# Example for macOS/Linux: excel_file_path <- "~/Desktop/Compiled_done.xlsx"
excel_file_path <- "path/to/your/Compiled_done.xlsx" # <<<--- REPLACE THIS WITH YOUR ACTUAL FILE PATH

# List all the sheet names exactly as they appear in your Excel file
sheet_names <- c(
  "CMAUP 2.0",
  "TCM-ID",
  "ITCM",
  "TCMBank",
  "TM-MC 2.0",
  "DCABM-TCM",
  "ETCM 1.0",
  "SymMap",
  "HERB 1.0",
  "HERB 2.0",
  "BATMAN-TCM"
)
# If your sheet names are different, please update this vector accordingly.

# Run the function to create "Combined data.xlsx"
message("\n--- Starting process to create 'Combined data.xlsx' ---")
combined_results <- create_combined_data_excel_r(excel_file_path, sheet_names, "Combined data.xlsx")

# Optional: Print a summary of the number of unique elements found for each column
message("\n--- Summary of Unique Elements Across All Sheets (per Column) ---")
if (length(combined_results) > 0) {
  for (col_name in names(combined_results)) {
    cat(paste0("Column '", col_name, "': ", nrow(combined_results[[col_name]]), " unique elements found across all sheets.\n"))
  }
} else {
  message("No unique data found or no files processed.")
}

Total.xlsx

# Load necessary libraries
# If you don't have these installed, run:
# install.packages("readxl")   # For read_excel
# install.packages("dplyr")    # For data manipulation
# install.packages("tidyr")    # Explicitly for pivot_wider
# install.packages("writexl")  # For saving output to Excel

library(readxl)
library(dplyr)
library(tidyr) # Explicitly load tidyr for pivot_wider
library(writexl)

# --- Function to combine all unique data into a single presence-absence matrix ---
create_total_combined_excel_r <- function(input_excel_file_path, sheet_names_list, output_file_name = "Total.xlsx") {
  
  if (!file.exists(input_excel_file_path)) {
    message(paste0("Error: Input Excel file not found at ", input_excel_file_path))
    return(NULL)
  }

  # Define the 5 specific column headers you are interested in
  target_columns <- c("Herbs", "Ingredients", "Formulas", "Targets", "Diseases")

  # List to store all unique, cleaned values from all target columns across all sheets
  all_unique_elements_across_all_categories <- character(0)

  # List to store processed data for each sheet, for easier lookup later
  processed_sheet_data <- list()

  message(paste0("Step 1: Collecting all unique elements and processing sheets from '", basename(input_excel_file_path), "'..."))

  for (sheet_name in sheet_names_list) {
    message(paste0("  Loading sheet: ", sheet_name))

    tryCatch({
      df <- read_excel(input_excel_file_path, sheet = sheet_name, col_names = TRUE)
      
      # Store unique values for each target column in the current sheet
      current_sheet_unique_values <- list()

      for (col in target_columns) {
        if (col %in% names(df)) {
          processed_col_values <- tolower(trimws(as.character(df[[col]])))
          unique_vals <- unique(na.omit(processed_col_values[
            processed_col_values != "" & !is.na(processed_col_values) & processed_col_values != "na" & processed_col_values != "nan"
          ]))
          
          # Add to the master list of all unique elements
          all_unique_elements_across_all_categories <- unique(c(all_unique_elements_across_all_categories, unique_vals))
          
          # Store for this sheet's lookup
          current_sheet_unique_values[[col]] <- unique_vals
        } else {
          message(paste0("    Warning: Column '", col, "' not found in sheet '", sheet_name, "'."))
          current_sheet_unique_values[[col]] <- character(0) # Ensure it's empty if column is missing
        }
      }
      processed_sheet_data[[sheet_name]] <- current_sheet_unique_values

    }, error = function(e) {
      message(paste0("Error loading or processing sheet '", sheet_name, "': ", e$message))
    })
  }

  # Sort the master list of all unique elements alphabetically
  all_unique_elements_across_all_categories <- sort(all_unique_elements_across_all_categories)

  message("\nStep 2: Constructing the combined presence-absence matrix...")

  # Create the base data frame for the combined output
  # Rows are all unique elements, columns are sheet names, initialized with 0s
  combined_df <- data.frame(Value = all_unique_elements_across_all_categories, stringsAsFactors = FALSE)
  
  # Add columns for each sheet, initialized to 0
  for (sheet_name in sheet_names_list) {
    combined_df[[sheet_name]] <- 0L # Use 0L for integer 0
  }

  # Populate the matrix with 1s where elements are present in a sheet
  if (nrow(combined_df) > 0) {
    for (sheet_name in sheet_names_list) {
      if (sheet_name %in% names(processed_sheet_data)) {
        # Get all unique elements found in *any* relevant column for this specific sheet
        elements_in_this_sheet <- unlist(processed_sheet_data[[sheet_name]])
        elements_in_this_sheet <- unique(elements_in_this_sheet) # Deduplicate within sheet's collected elements

        # Mark 1 for presence
        combined_df[[sheet_name]][combined_df$Value %in% elements_in_this_sheet] <- 1L
      }
    }
  } else {
    message("No unique elements found across all sheets to combine.")
  }

  # --- Save the results to the new Excel file ---
  output_list_for_excel <- list("All_Combined_Elements" = combined_df) # Name the single sheet
  
  if (nrow(combined_df) > 0) {
    writexl::write_xlsx(output_list_for_excel, output_file_name)
    message(paste0("\nCombined data (all unique elements with sheet occurrences) saved to '", output_file_name, "'"))
  } else {
    message(paste0("\nNo data to save to '", output_file_name, "'. This might mean no unique elements were found across all categories/sheets."))
  }

  return(combined_df) # Return the combined data frame for potential further use
}

# --- How to run the analysis on your machine ---

# IMPORTANT: Set this path to your "Compiled_done_2.xlsx" file
# Example for Windows: excel_file_path <- "C:/Users/YourUsername/Desktop/Compiled_done_2.xlsx"
# Example for macOS/Linux: excel_file_path <- "~/Desktop/Compiled_done_2.xlsx"
excel_file_path <- "path/to/your/Compiled_done_2.xlsx" # <<<--- REPLACE THIS WITH YOUR ACTUAL FILE PATH

# List all the sheet names exactly as they appear in your Excel file
sheet_names <- c(
  "CMAUP 2.0",
  "TCM-ID",
  "ITCM",
  "TCMBank",
  "TM-MC 2.0",
  "DCABM-TCM",
  "ETCM 1.0",
  "SymMap",
  "HERB 1.0",
  "HERB 2.0",
  "BATMAN-TCM"
)
# If your sheet names are different, please update this vector accordingly.

# Run the function to create "Total.xlsx"
message("\n--- Starting process to create 'Total.xlsx' ---")
total_combined_results <- create_total_combined_excel_r(excel_file_path, sheet_names, "Total.xlsx")

# Optional: Print a summary of the number of unique elements found
message("\n--- Summary of All Unique Elements Across All Categories and Sheets ---")
if (!is.null(total_combined_results) && nrow(total_combined_results) > 0) {
  cat(paste0("Total unique elements found across all categories and sheets: ", nrow(total_combined_results), "\n"))
} else {
  message("No unique elements found or processing failed.")
}



Total output + MDS:

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
# install.packages("cluster") # For divisive hierarchical clustering (diana function)
# install.packages("factoextra") # For enhanced dendrogram plotting, and fviz_pca/fviz_mca
# install.packages("ggplot2") # For custom MDS plot
# install.packages("vegan") # For Jaccard distance (vegdist)

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
# library(FactoMineR) # Removed: Load the FactoMineR package for MCA
library(ggplot2) # Load ggplot2 for MDS plotting
library(vegan) # Load vegan for vegdist function

# --- Function to process individual Excel sheets into a binary format ---
# This function is crucial for preparing data for MCA, Jaccard, and Overlap analyses.
process_excel_sheet <- function(df, sheet_name) {
  # Drop columns that are entirely NA
  df <- df %>% select(where(~!all(is.na(.))))

  # Filter out rows where 'Value' is NA or empty string
  df <- df %>% filter(!is.na(Value) & Value != "")

  if (nrow(df) == 0) {
    message(paste0("  Sheet '", sheet_name, "' has no valid rows after initial filtering. Skipping."))
    return(NULL)
  }

  # Ensure 'df' is a tibble for robust dplyr operations
  df <- as_tibble(df)

  # Identify data source columns (all columns except 'Value')
  data_source_cols <- setdiff(colnames(df), "Value")

  if (length(data_source_cols) == 0) {
    message(paste0("  Sheet '", sheet_name, "' has no data source columns besides 'Value' after cleaning. Skipping binarization.")) # Removed MCA mention
    return(NULL)
  }

  # Ensure all other columns (data sources) are treated as binary (1s or 0s)
  # Convert any non-NA, non-zero values to 1, and NA/0 to 0
  df_binary <- df %>%
    mutate(across(all_of(data_source_cols), ~ case_when( # Use all_of() for robust column selection
      !is.na(.) & . != 0 ~ 1L, # Non-NA and non-zero -> 1 (integer)
      TRUE ~ 0L # Otherwise (NA or zero) -> 0 (integer)
    ))) %>%
    distinct() # Remove duplicate rows that might arise from binarization

  if (nrow(df_binary) == 0) {
    message(paste0("  Sheet '", sheet_name, "' has no rows with present elements after binarization. Skipping."))
    return(NULL)
  }

  return(df_binary)
}


# --- Function to perform Multidimensional Scaling (MDS) ---
perform_mds_analysis <- function(jaccard_distance_matrix, category_name, output_dir_mds) {
  
  # Ensure the MDS output directory exists
  if (!dir.exists(output_dir_mds)) {
    dir.create(output_dir_mds, recursive = TRUE)
  }

  # Robust input validation: Check if it's a valid 'dist' object and not NULL
  if (is.null(jaccard_distance_matrix) || !inherits(jaccard_distance_matrix, "dist")) {
    message(paste0("  Skipping MDS for category '", category_name, "': Invalid or NULL distance matrix provided."))
    return(NULL)
  }

  # Get the number of items from the distance matrix (correct way for 'dist' objects)
  num_items <- attr(jaccard_distance_matrix, "Size")

  # Check if the number of items is valid for MDS (at least 2)
  if (is.null(num_items) || num_items < 2) {
    message(paste0("  Skipping MDS for category '", category_name, "': Insufficient items (", ifelse(is.null(num_items), "NULL", num_items), "). Needs at least 2 for MDS."))
    return(NULL)
  }
  
  # Ensure the distance matrix does not contain NA or Inf values that could cause issues
  if (any(is.na(jaccard_distance_matrix)) || any(is.infinite(jaccard_distance_matrix))) {
      message(paste0("  Skipping MDS for category '", category_name, "': Distance matrix contains NA or Inf values."))
      return(NULL)
  }

  # Check for constant distances (no variance) after ensuring no NAs/Infs
  if (sd(as.vector(jaccard_distance_matrix), na.rm = TRUE) == 0) {
    message(paste0("  Skipping MDS for category '", category_name, "': All distances are identical. No variance to project."))
    return(NULL)
  }

  message(paste0("  Performing MDS for category: ", category_name))

  tryCatch({
    # Perform Classical (Metric) MDS
    # k=2 for 2-dimensional representation
    mds_result <- cmdscale(jaccard_distance_matrix, k = 2, eig = TRUE)

    # Check if MDS returned valid dimensions (points can be NULL or have fewer than k dimensions if data is bad)
    if (is.null(mds_result$points) || ncol(mds_result$points) < 2) {
        message(paste0("  MDS for '", category_name, "' did not return enough dimensions for plotting (expected 2, got ", ncol(mds_result$points), "). Skipping MDS plot."))
        return(NULL)
    }

    # Extract points for plotting
    mds_points <- as.data.frame(mds_result$points)
    colnames(mds_points) <- c("Dim1", "Dim2")
    # Use rownames from the MDS result itself, which should be inherited from the distance matrix
    mds_points$Source <- rownames(mds_result$points)

    # Calculate stress (goodness of fit)
    original_dist_vec <- as.vector(jaccard_distance_matrix)
    mds_dist_vec <- as.vector(dist(mds_points[, c("Dim1", "Dim2")]))
    
    # Filter out NA values for stress calculation from both vectors
    valid_indices <- !is.na(original_dist_vec) & !is.na(mds_dist_vec)
    original_dist_vec <- original_dist_vec[valid_indices]
    mds_dist_vec <- mds_dist_vec[valid_indices]

    # Handle cases where all filtered distances are zero to avoid division by zero
    if (sum(original_dist_vec^2) == 0) {
      stress <- 0 # Perfect fit if all distances are zero
      stress_label <- "Stress-1: 0 (All distances are zero)"
    } else {
      stress <- sqrt(sum((original_dist_vec - mds_dist_vec)^2) / sum(original_dist_vec^2))
      stress_label <- paste0("Stress-1: ", round(stress, 3))
    }

    # Plot MDS results using ggplot2
    pdf_mds <- file.path(output_dir_mds, paste0("MDS_Plot_", gsub(" ", "_", category_name), ".pdf"))
    pdf(pdf_mds, width = 10, height = 10)
    
    p <- ggplot(mds_points, aes(x = Dim1, y = Dim2, label = Source)) +
      geom_point(size = 3, color = "steelblue") +
      geom_text(vjust = -1, hjust = 0.5, size = 3) +
      geom_hline(yintercept = 0, linetype = "dashed", color = "gray") +
      geom_vline(xintercept = 0, linetype = "dashed", color = "gray") +
      labs(title = paste("MDS Plot for Database Similarity (", category_name, ")"),
           subtitle = stress_label,
           x = "Dimension 1",
           y = "Dimension 2") +
      theme_minimal() +
      theme(plot.title = element_text(hjust = 0.5, face = "bold"),
            plot.subtitle = element_text(hjust = 0.5))
    
    print(p)
    dev.off()
    message(paste0("  Saved MDS plot for '", category_name, "' to '", pdf_mds, "'"))

  }, error = function(e) {
    message(paste0("  ERROR performing MDS for '", category_name, "': ", e$message))
    message("  This might happen if the distance matrix is problematic (e.g., contains NaNs or Inf) or if cmdscale fails.")
  }, finally = {
    if (dev.cur() != 1) { dev.off() } # Ensure graphics device is closed if an error occurs
  })
}


# --- Function to calculate and plot Jaccard Similarity and Overlap Counts Matrices ---
calculate_and_plot_matrices <- function(combined_data_list, output_dir = "./Output_Matrices/") {

  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
  }

  for (col_name in names(combined_data_list)) { # col_name will now be "All_Combined_Elements"
    df_binary <- combined_data_list[[col_name]]

    if (is.null(df_binary) || nrow(df_binary) == 0 || ncol(df_binary) < 3) {
      message(paste0("Skipping matrix analysis for category '", col_name, "': No data or insufficient data source columns (need at least 2 besides 'Value')."))
      next
    }

    binary_matrix <- as.matrix(df_binary[, -1]) # Exclude the 'Value' column

    if (ncol(binary_matrix) < 2) {
      message(paste0("Skipping matrix analysis for category '", col_name, "': Only one data source column found after removing 'Value'. Need at least 2 for comparisons."))
      next
    }

    if (all(binary_matrix == 0)) {
      message(paste0("Skipping matrix analysis for category '", col_name, "': All elements are 0 across all sheets. No overlaps or unique elements to calculate."))
      next
    }

    # --- Jaccard Similarity Matrix ---
    jaccard_dist_matrix <- vegdist(t(binary_matrix), method = "jaccard") # Transpose to calculate distance between columns (sources)
    jaccard_sim_matrix <- 1 - as.matrix(jaccard_dist_matrix)
    diag(jaccard_sim_matrix) <- 1 # Diagonal should be 1 (self-similarity)
    jaccard_sim_matrix[is.nan(jaccard_sim_matrix)] <- 0 # Handle cases where sets are empty, Jaccard is undefined

    # Create a version of the matrix specifically for plotting with more decimal points
    # This ensures the numbers displayed on the PDF plot retain the desired precision.
    # The original 'jaccard_sim_matrix' (full precision) is still saved to CSV.
    jaccard_sim_matrix_for_plot <- round(jaccard_sim_matrix, 4) # Round to 4 decimal places for display

    csv_file_name_jaccard <- file.path(output_dir, paste0("Jaccard_Similarity_Matrix_", gsub(" ", "_", col_name), ".csv"))
    write.csv(jaccard_sim_matrix, csv_file_name_jaccard, row.names = TRUE) # This writes the *unrounded* matrix
    message(paste0("Saved Jaccard similarity matrix for '", col_name, "' to '", csv_file_name_jaccard, "'"))

    pdf_file_name_jaccard <- file.path(output_dir, paste0("Jaccard_Similarity_Matrix_", gsub(" ", "_", col_name), ".pdf"))
    
    message(paste0("  DEBUG: Plotting '", col_name, "' Jaccard Similarity Matrix. Dimensions: ",
                   paste(dim(jaccard_sim_matrix_for_plot), collapse = "x"), # Using _for_plot for debug message
                   ", Range: [", round(min(jaccard_sim_matrix_for_plot, na.rm = TRUE), 4), ", ", round(max(jaccard_sim_matrix_for_plot, na.rm = TRUE), 4), "]",
                   ", All zeros: ", all(jaccard_sim_matrix_for_plot == 0)))

    tryCatch({
      pdf(pdf_file_name_jaccard, width = 10, height = 10)
      
      if (nrow(jaccard_sim_matrix_for_plot) >= 2 && ncol(jaccard_sim_matrix_for_plot) >= 2) {
        corrplot(jaccard_sim_matrix_for_plot, # Use the rounded matrix for plotting
                 method = "color",
                 type = "full",
                 order = "hclust",
                 tl.col = "black",
                 tl.srt = 45,
                 addCoef.col = "black",
                 number.cex = 1.2, # Further increased this value for better visibility of decimal points
                 main = paste("Jaccard Similarity Matrix for", gsub("_", " ", col_name)), # Adjust title for readability
                 col = colorRampPalette(c("white", "lightblue", "darkblue"))(100),
                 cl.lim = c(0, 1) # Explicitly set color legend limits for Jaccard (0 to 1)
        )
      } else {
        grid.newpage()
        grid.text(paste("Jaccard Similarity Matrix for", gsub("_", " ", col_name), "\n(Not suitable for heatmap, showing as table)"),
                  y = 0.95, gp = gpar(fontsize = 12, fontface = "bold"))
        
        display_data <- round(jaccard_sim_matrix_for_plot, 4) # Ensure table also uses rounded data
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

    # --- Overlap Counts Matrix (CSV and Table PDF) ---
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
      grid.text(paste("Pairwise Intersection Counts for", gsub("_", " ", col_name)), # Adjust title for readability
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

    # --- NEW: Overlap Counts Matrix (Scaled Heatmap PDF) ---
    pdf_file_name_overlap_counts_scaled <- file.path(output_dir, paste0("Overlap_Counts_Matrix_Scaled_", gsub(" ", "_", col_name), ".pdf"))

    # Scale the overlap counts matrix for plotting colors (e.g., 0 to 1)
    # Divide by the maximum value to get values between 0 and 1
    max_overlap_count <- max(overlap_counts_matrix, na.rm = TRUE)
    if (max_overlap_count > 0) {
      scaled_overlap_counts_for_plot <- overlap_counts_matrix / max_overlap_count
    } else {
      # If max_overlap_count is 0, set all to 0 to avoid division by zero
      scaled_overlap_counts_for_plot <- matrix(0, nrow = nrow(overlap_counts_matrix), ncol = ncol(overlap_counts_matrix),
                                               dimnames = dimnames(overlap_counts_matrix))
    }
    
    # Ensure values are rounded for display on the plot coefficients
    scaled_overlap_counts_for_plot_rounded <- round(scaled_overlap_counts_for_plot, 4)

    message(paste0("  DEBUG: Plotting '", col_name, "' Scaled Overlap Counts Matrix. Dimensions: ",
                   paste(dim(scaled_overlap_counts_for_plot_rounded), collapse = "x"),
                   ", Range: [", round(min(scaled_overlap_counts_for_plot_rounded, na.rm = TRUE), 4), ", ", round(max(scaled_overlap_counts_for_plot_rounded, na.rm = TRUE), 4), "]",
                   ", All zeros: ", all(scaled_overlap_counts_for_plot_rounded == 0)))

    tryCatch({
      pdf(pdf_file_name_overlap_counts_scaled, width = 10, height = 10)

      if (nrow(scaled_overlap_counts_for_plot_rounded) >= 2 && ncol(scaled_overlap_counts_for_plot_rounded) >= 2) {
        corrplot(scaled_overlap_counts_for_plot_rounded,
                 method = "color",
                 type = "full",
                 order = "hclust",
                 tl.col = "black",
                 tl.srt = 45,
                 addCoef.col = "black",
                 number.cex = 0.8, # Adjust as needed for visibility
                 main = paste("Scaled Overlap Counts Matrix for", gsub("_", " ", col_name)),
                 col = colorRampPalette(c("white", "lightcoral", "darkred"))(100),
                 cl.lim = c(0, 1) # Color legend from 0 to 1 as it's scaled
        )
      } else {
        grid.newpage()
        grid.text(paste("Scaled Overlap Counts Matrix for", gsub("_", " ", col_name), "\n(Not suitable for heatmap, showing as table)"),
                  y = 0.95, gp = gpar(fontsize = 12, fontface = "bold"))
        
        display_data <- round(scaled_overlap_counts_for_plot_rounded, 4)
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
      message(paste0("Successfully generated plot for '", col_name, "' to '", pdf_file_name_overlap_counts_scaled, "'"))
    }, error = function(e) {
      message(paste0("ERROR: Failed to generate PDF for '", col_name, "' (Scaled Overlap Counts). Error: ", e$message))
    }, finally = {
      if (dev.cur() != 1) {
        dev.off()
      }
    })

    # --- Overlap Percentage Matrix ---
    num_sheets <- nrow(overlap_counts_matrix)
    overlap_percentage_matrix <- matrix(0, nrow = num_sheets, ncol = num_sheets,
                                         dimnames = dimnames(overlap_counts_matrix))

    for (i in 1:num_sheets) {
      for (j in 1:num_sheets) {
        denominator_size <- overlap_counts_matrix[j,j] # Size of the column (sheet) j
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
    
    scaled_overlap_percentage_matrix_for_plot <- overlap_percentage_matrix / 100 # Scale 0-1 for plotting

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
                 main = paste("Overlap Percentage (Intersection / Column Set Size) Matrix for", gsub("_", " ", col_name), " (Scaled 0-1)"), # Adjust title
                 col = colorRampPalette(c("white", "lightgreen", "darkgreen"))(100),
                 cl.lim = c(0, 1),
                 cl.pos = "r"
        )
      } else {
        grid.newpage()
        grid.text(paste("Overlap Percentage Matrix for", gsub("_", " ", col_name), "\n(Not suitable for heatmap, showing as table)"),
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
} # End of calculate_and_plot_matrices function

# --- Function to perform Hierarchical Clustering ---
perform_hierarchical_clustering <- function(jaccard_matrix_path, output_dir) {
  
  # Ensure the output directory exists
  if (!dir.exists(output_dir)) {
    dir.create(output_dir, recursive = TRUE)
  }

  # Extract category name from file path for titles
  file_name <- basename(jaccard_matrix_path)
  category_name <- gsub("Jaccard_Similarity_Matrix_|.csv", "", file_name)
  plot_title_suffix <- paste("Databases based on", gsub("_", " ", category_name), "Jaccard Distance")

  # 1. Load the Jaccard Similarity Matrix
  if (!file.exists(jaccard_matrix_path)) {
    message(paste0("Skipping clustering for '", category_name, "': Jaccard matrix file not found at '", jaccard_matrix_path, "'"))
    return(NULL)
  }
  similarity_matrix <- read.csv(jaccard_matrix_path, row.names = 1)
  similarity_matrix <- as.matrix(similarity_matrix)

  # Check for valid dimensions for clustering (at least 2x2)
  if (nrow(similarity_matrix) < 2 || ncol(similarity_matrix) < 2) {
    message(paste0("Skipping clustering for '", category_name, "': Matrix is too small (", nrow(similarity_matrix), "x", ncol(similarity_matrix), "). Needs at least 2x2 for clustering."))
    return(NULL)
  }
  
  # Ensure all values are numeric
  if (!all(sapply(similarity_matrix, is.numeric))) {
    message(paste0("Skipping clustering for '", category_name, "': Matrix contains non-numeric values after loading. Please check the CSV file."))
    return(NULL)
  }

  # Convert Jaccard Similarity to Distance
  distance_matrix <- as.dist(1 - similarity_matrix)

  # Check if all distances are identical (e.g., all 0s or all 1s)
  if (sd(as.vector(distance_matrix), na.rm = TRUE) == 0) {
    message(paste0("Skipping clustering for '", category_name, "': All distance values are identical. No meaningful clustering can be performed."))
    return(NULL)
  }

  message(paste0("\n--- Performing Hierarchical Clustering for ", category_name, " ---"))

  # --- Agglomerative Hierarchical Clustering (agnes) ---
  pdf_agnes <- file.path(output_dir, paste0("Clustering_Agglomerative_Dendrogram_", gsub(" ", "_", category_name), ".pdf"))
  tryCatch({
    agnes_cluster <- agnes(distance_matrix, method = "ward") # Ward's method is common
    pdf(pdf_agnes, width = 12, height = 8)
    print(fviz_dend(agnes_cluster,
                    k = NULL, # No specific number of clusters
                    cex = 0.8,
                    main = paste("Agglomerative Hierarchical Clustering (Ward's Method) for", plot_title_suffix),
                    horiz = FALSE, # Vertical dendrogram
                    rect = FALSE, # Do not draw rectangles around clusters by default
                    ggtheme = theme_minimal()
    ))
    dev.off()
    message(paste0("  Saved Agglomerative Dendrogram for '", category_name, "' to '", pdf_agnes, "'"))
  }, error = function(e) {
    message(paste0("  ERROR generating Agglomerative Dendrogram for '", category_name, "': ", e$message))
  }, finally = {
    if (dev.cur() != 1) { dev.off() }
  })

  # --- Divisive Hierarchical Clustering (diana) ---
  pdf_diana <- file.path(output_dir, paste0("Clustering_Divisive_Dendrogram_", gsub(" ", "_", category_name), ".pdf"))
  tryCatch({
    diana_cluster <- diana(distance_matrix)
    pdf(pdf_diana, width = 12, height = 8)
    print(fviz_dend(diana_cluster,
                    k = NULL, # No specific number of clusters
                    cex = 0.8,
                    main = paste("Divisive Hierarchical Clustering for", plot_title_suffix),
                    horiz = FALSE, # Vertical dendrogram
                    rect = FALSE, # Do not draw rectangles around clusters by default
                    ggtheme = theme_minimal()
    ))
    dev.off()
    message(paste0("  Saved Divisive Dendrogram for '", category_name, "' to '", pdf_diana, "'"))
  }, error = function(e) {
    message(paste0("  ERROR generating Divisive Dendrogram for '", category_name, "': ", e$message))
  }, finally = {
    if (dev.cur() != 1) { dev.off() }
  })
}


# --- Main execution script ---

# --- Part 1: Process "Total.xlsx" for Jaccard, Overlap, Clustering, and MDS ---

# IMPORTANT: Set this path to your "Total.xlsx" file
# Ensure "Total.xlsx" has been generated by a previous step!
# Example for Windows: input_excel_file_total <- "C:/Users/YourUsername/Desktop/Total.xlsx"
# Example for macOS/Linux: input_excel_file_total <- "~/Desktop/Total.xlsx"
input_excel_file_total <- "path/to/your/Total.xlsx" # <<<--- REPLACE THIS WITH YOUR ACTUAL FILE PATH

# Define the single sheet name in "Total.xlsx" (assuming it contains all combined elements)
total_sheet_name <- "All_Combined_Elements"  

# Define the new output directory for all matrices and plots related to "Total.xlsx"
output_dir_total <- "./Total_output_Matrices/"

message("\n--- Starting process to load 'Total.xlsx' ---")

if (!file.exists(input_excel_file_total)) {
  stop(paste0("Error: 'Total.xlsx' not found at ", input_excel_file_total, ". Please ensure the file exists and the path is correct, and that it was generated by the previous step."))
}

# Load and process the single sheet from "Total.xlsx"
processed_total_sheet <- NULL
message(paste0("  Loading sheet: ", total_sheet_name, " from '", basename(input_excel_file_total), "'..."))
tryCatch({
  df_total <- read_excel(input_excel_file_total, sheet = total_sheet_name, col_names = TRUE)
  processed_total_sheet <- process_excel_sheet(df_total, total_sheet_name)
}, error = function(e) {
  stop(paste0("Error loading or processing sheet '", total_sheet_name, "' from '", input_excel_file_total, "': ", e$message))
})


message("\n--- Summary of Combined Data Loaded from 'Total.xlsx' ---")
if (!is.null(processed_total_sheet) && nrow(processed_total_sheet) > 0) {
  cat(paste0("Sheet '", total_sheet_name, "': ", nrow(processed_total_sheet), " unique elements processed.\n"))
} else {
  message("No valid data found or loaded from 'Total.xlsx' after processing. Skipping analyses.")
}

# Only proceed with analyses if processed_total_sheet contains valid data
if (!is.null(processed_total_sheet) && nrow(processed_total_sheet) > 0) {
    # --- Calculate and Plot Jaccard Similarity and Overlap Counts Matrices ---
    message("\n--- Calculating and Plotting Jaccard Similarity and Overlap Counts Matrices for Total Combined Data ---")
    
    # The calculate_and_plot_matrices function expects a list of dataframes
    # So, we pass processed_total_sheet wrapped in a list.
    calculate_and_plot_matrices(list("All_Combined_Elements" = processed_total_sheet), output_dir = output_dir_total)

    # --- Perform and plot both Agglomerative and Divisive Hierarchical Clustering ---
    message("\n--- Starting Hierarchical Clustering (Agglomerative & Divisive) Analysis for Total Combined Data ---")

    # The Jaccard matrix for the total combined data will be in the new output directory
    jaccard_matrix_path_total <- file.path(output_dir_total, paste0("Jaccard_Similarity_Matrix_", gsub(" ", "_", total_sheet_name), ".csv"))
    perform_hierarchical_clustering(jaccard_matrix_path_total, output_dir = output_dir_total)

    # --- Perform Multidimensional Scaling (MDS) ---
    # Need to regenerate the Jaccard distance matrix object for MDS, as the function expects a 'dist' object
    # (The clustering function loaded the CSV, but MDS needs the original 'dist' object for cmdscale)
    message("\n--- Starting Multidimensional Scaling (MDS) Analysis for Total Combined Data ---")
    binary_matrix_for_mds <- as.matrix(processed_total_sheet %>% select(-Value))
    if (ncol(binary_matrix_for_mds) >= 2) {
        jaccard_distance_for_mds <- vegdist(t(binary_matrix_for_mds), method = "jaccard")
        perform_mds_analysis(jaccard_distance_for_mds, total_sheet_name, output_dir_total)
    } else {
        message(paste0("  Skipping MDS for '", total_sheet_name, "': Insufficient data source columns (", ncol(binary_matrix_for_mds), "). Need at least 2 for MDS."))
    }

} else {
  message("\nSkipping all further analyses for 'Total.xlsx' due to no valid processed data.")
}

message("\n--- All analyses complete for Total.xlsx. Check the './Total_output_Matrices/' folder for results. ---")

