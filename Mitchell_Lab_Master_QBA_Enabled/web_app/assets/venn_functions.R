knitr::opts_chunk$set(echo = TRUE)
# Function to install and load necessary packages
install_if_missing <- function(packages) {
  for (package in packages) {
    if (!require(package, character.only = TRUE)) {
      install.packages(package, dependencies = TRUE)
      library(package, character.only = TRUE)
    }
  }
}

# List of required packages
required_packages <- c("readxl", "dplyr", "devtools", "VennDiagram", "ggplot2", "openxlsx")

# Install and load required packages
install_if_missing(required_packages)

# Load necessary libraries
library(readxl)
library(dplyr)
library(devtools)
library(VennDiagram)
library(ggplot2)
library(grid) 
# Normalize column names so matching is case-insensitive and punctuation-insensitive
clean_col_name <- function(x) {
  gsub("[^a-z0-9]", "", tolower(x))
}

# Detect CpG column in multi-column sheets
find_cpg_column <- function(df, path, sheet) {
  nm <- names(df)
  if (nzchar(cfg$id_column)) {
    hit <- which(tolower(nm) == tolower(cfg$id_column))
    if (length(hit) != 1) stop("Requested ID column not found uniquely: ", cfg$id_column)
    return(hit)
  }
  nm_clean <- clean_col_name(nm)

  # Prefer exact cleaned match: CpG, cpg, CPG, c_p_g, etc.
  exact <- which(nm_clean == "cpg")

  if (length(exact) == 1) {
    return(exact)
  }

  if (length(exact) > 1) {
    stop(
      "Multiple exact CpG columns detected in sheet '", sheet,
      "' of file ", basename(path), ": ",
      paste(nm[exact], collapse = ", ")
    )
  }

  # Then allow partial matches like CpG_ID, cpgName, CpG.Site, etc.
  partial <- which(grepl("cpg", nm_clean))

  if (length(partial) == 1) {
    return(partial)
  }

  if (length(partial) > 1) {
    stop(
      "Multiple possible CpG columns detected in sheet '", sheet,
      "' of file ", basename(path), ": ",
      paste(nm[partial], collapse = ", "),
      ". Please rename the intended column to 'CpG'."
    )
  }

  stop(
    "No CpG column detected in multi-column sheet '", sheet,
    "' of file ", basename(path), "."
  )
}

# Detect optional Gene Name column
find_gene_name_column <- function(df) {
  nm <- names(df)
  if (nzchar(cfg$gene_column)) {
    hit <- which(tolower(nm) == tolower(cfg$gene_column))
    if (length(hit) != 1) stop("Requested gene column not found uniquely: ", cfg$gene_column)
    return(hit)
  }
  nm_clean <- clean_col_name(nm)

  # Catches Gene Name, GeneName, gene_name, GeneName_NCBI, etc.
  matches <- which(grepl("genename", nm_clean))

  if (length(matches) == 0) {
    return(NA_integer_)
  }

  exact <- which(nm_clean == "genename")

  if (length(exact) >= 1) {
    return(exact[1])
  }

  if (length(matches) > 1) {
    message(
      "Multiple Gene Name-like columns detected: ",
      paste(nm[matches], collapse = ", "),
      ". Using: ", nm[matches[1]]
    )
  }

  matches[1]
}

# Helper Function to read a single sheet from an Excel file
read_set_from_excel_sheet <- function(path, sheet) {

  # Read everything without headers first.
  # This protects true single-column unlabeled files from losing row 1.
  df_raw <- readxl::read_excel(
    path,
    sheet = sheet,
    col_names = FALSE,
    .name_repair = "minimal"
  )

  if (ncol(df_raw) < 1) {
    stop("Sheet '", sheet, "' in file ", basename(path), " has no columns.")
  }

  # CASE 1: single-column sheet
  # Assume the column is CpG. If it has a header named CpG, remove that header row.
  if (ncol(df_raw) == 1) {
    cpg_raw <- trimws(as.character(df_raw[[1]]))
    cpg_raw <- cpg_raw[!is.na(cpg_raw) & cpg_raw != ""]

    # Remove a literal CpG header if present
    cpg_raw <- cpg_raw[clean_col_name(cpg_raw) != "cpg" & (!nzchar(cfg$id_column) | tolower(cpg_raw) != tolower(cfg$id_column))]

    v <- unique(cpg_raw)

    attr(v, "id_col") <- "Column 1 assumed as CpG"
    attr(v, "gene_map") <- NULL

    return(v)
  }

  # CASE 2: multi-column sheet
  # Determine whether the first row contains headers.
  first_row <- as.character(unlist(df_raw[1, ], use.names = FALSE))
  first_row_clean <- clean_col_name(first_row)

  has_cpg_header <- any(first_row_clean == "cpg" | grepl("cpg", first_row_clean))
  has_gene_header <- any(grepl("genename", first_row_clean))

  if (has_cpg_header || has_gene_header || (nzchar(cfg$id_column) && any(tolower(first_row) == tolower(cfg$id_column), na.rm=TRUE))) {
    # Treat row 1 as the header row
    header <- first_row

    # Replace blank/NA headers with generic names
    header[is.na(header) | trimws(header) == ""] <- paste0("Column_", which(is.na(header) | trimws(header) == ""))

    df <- df_raw[-1, , drop = FALSE]
    names(df) <- header
  } else {
    # Fallback: no recognizable header in row 1
    # This should be rare for multi-column files, but gives a clearer error.
    names(df_raw) <- paste0("Column_", seq_len(ncol(df_raw)))
    df <- df_raw

    stop(
      "No header row containing a CpG column was detected in multi-column sheet '",
      sheet, "' of file ", basename(path), 
      "'. First row values were: ",
      paste(first_row, collapse = " | ")
    )
  }

  cpg_col <- find_cpg_column(df, path, sheet)

  cpg_raw <- trimws(as.character(df[[cpg_col]]))
  keep <- !is.na(cpg_raw) & cpg_raw != ""

  v <- unique(cpg_raw[keep])

  attr(v, "id_col") <- names(df)[cpg_col]

  # Optional Gene Name metadata for Excel outputs only
  gene_col <- find_gene_name_column(df)

  if (!is.na(gene_col)) {
    gene_raw <- trimws(as.character(df[[gene_col]]))
    gene_raw[is.na(gene_raw) | toupper(gene_raw) == "NA"] <- ""

    meta <- data.frame(
      CpG = cpg_raw,
      GeneName = gene_raw,
      stringsAsFactors = FALSE
    )

    meta <- meta[!is.na(meta$CpG) & meta$CpG != "", , drop = FALSE]

    meta <- meta |>
      dplyr::group_by(CpG) |>
      dplyr::summarise(
        GeneName = paste(unique(GeneName[GeneName != ""]), collapse = "; "),
        .groups = "drop"
      )

    attr(v, "gene_map") <- meta
    attr(v, "gene_col") <- names(df)[gene_col]
  } else {
    attr(v, "gene_map") <- NULL
  }

  v
}

# Helper Function to read one or more sets from an Excel file
# - If the workbook has one sheet, use the FILE name as the set/group name
# - If the workbook has multiple sheets, use each SHEET name as a set/group name
read_sets_from_excel <- function(path) {
  sheets <- readxl::excel_sheets(path)

  if (length(sheets) < 1) {
    stop("File ", basename(path), " has no sheets.")
  }

  out <- list()

  if (length(sheets) == 1) {
    # Single-sheet workbook: retain the file name as the group/set title
    nm_raw <- sub("\\.[^.]*$", "", basename(path))

    out[[nm_raw]] <- read_set_from_excel_sheet(path, sheet = sheets[1])

  } else {
    # Multi-sheet workbook: retain each sheet name as the group/set title
    for (sh in sheets) {
      out[[sh]] <- read_set_from_excel_sheet(path, sheet = sh)
    }
  }

  out
}

# Helper Function to select Excel files (multi-select on Windows)
select_excel_files <- function(caption = "Select Excel files") {
if (.Platform$OS.type == "windows") {
files <- utils::choose.files(caption = caption, multi = TRUE)
} else {
# Fallback: single-file chooser on non-Windows
f <- file.choose()
files <- if (nzchar(f)) f else character(0)
}
files
}

# Main loader to interactively load multiple Excel files
# and store them in a named list + (optionally) global variables
load_sets_interactively <- function() {
set_list <- list()
loaded_paths <- character(0)

repeat {
cat("Select one or more Excel files...\n")
files <- select_excel_files("Select Excel files")

if (length(files) == 0) {
  cat("No files selected.\n")
} else {
  # Separate new vs already-loaded files
  new_files <- setdiff(files, loaded_paths)
  dup_files <- intersect(files, loaded_paths)

  if (length(dup_files) > 0) {
    message("Files have already been read: ",
            paste(basename(dup_files), collapse = ", "))
  }

 for (f in new_files) {
  cat("Reading file:", basename(f), "...\n")

  sets_from_file <- read_sets_from_excel(f)

  for (nm_raw in names(sets_from_file)) {
    nm <- nm_raw

    # If the same group/set name already exists, make it unique
    # This prevents accidental overwriting if two files/sheets have the same name
    if (nm %in% names(set_list)) {
      nm <- make.unique(c(names(set_list), nm))[length(names(set_list)) + 1]
      message("Duplicate set name detected. Renaming to: ", nm)
    }

    cat("  Adding set:", nm, "\n")

    v <- sets_from_file[[nm_raw]]

    # Store in list using the visible group/set title
    set_list[[nm]] <- v

    # Optional: also create a safe global variable
    # This uses make.names() only for the R variable, NOT for the visible set title
    assign(make.names(nm), v, envir = .GlobalEnv)
  }
}

  loaded_paths <- union(loaded_paths, new_files)
}

# Ask if the user wants to add more files
ans <- readline("Would you like to upload more files? (y/n) ")
ans_trim <- tolower(trimws(ans))
if (!ans_trim %in% c("y", "yes", "yeah")) break

}

if (length(set_list) < 2) {
stop("Need at least two sets loaded to perform a Venn comparison.")
}

cat("Loaded sets:\n")
print(names(set_list))
set_list
}

# Run the interactive loader once; this creates set_list
set_list <- load_sets_interactively()
sort_gene_asc <- function(v) {
  nums <- suppressWarnings(as.integer(sub(".*?(\\d+)\\s*$", "\\1", v)))
  nums[is.na(nums)] <- Inf
  v[order(nums, v, decreasing = FALSE)]
}
# Helper function to display Venn diagram
display_venn <- function(x, ...){
  grid.newpage()
  venn_object <- venn.diagram(x, filename = NULL, ...)
  grid.draw(venn_object)
}
base_sets_in_label <- function(label) {
  parts <- strsplit(label, "∩", fixed = TRUE)[[1]]
  # Remove the trailing " Only" to get the canonical Set name
  sub("\\s+Only$", "", parts)
}
# Group genes by exclusive Venn intersections
venn_group_table <- function(x) {
  stopifnot(is.list(x), !is.null(names(x)))
  nm <- names(x)
  universe <- sort(unique(unlist(x, use.names = FALSE)))
  membership <- sapply(x, function(s) universe %in% s)   # rows=genes, cols=sets
 # raw labels like "Set 1∩Set 2" or "Set 3"
  raw_labels <- apply(membership, 1, function(row) paste(nm[row], collapse = "∩"))
  raw_labels[raw_labels == ""] <- "(none)"
# Rename singleton intersections "Set n" -> "Set n Only"
  labels <- vapply(strsplit(raw_labels, "∩", fixed = TRUE), function(parts) {
    if (length(parts) == 1 && parts != "(none)") paste0(parts, " Only") else paste(parts, collapse = "∩")
  }, character(1))
  
  tibble::tibble(gene = universe, intersect = labels) |>
    dplyr::group_by(intersect) |>
    dplyr::summarise(
      n = dplyr::n(),
      genes = list(gene),
      .groups = "drop"
    ) |>
    dplyr::arrange(dplyr::desc(n), intersect) |>
    dplyr::mutate(genes_csv = vapply(genes, function(g) paste(g, collapse = ", "), character(1)))
}

# Does any selected set contain Gene Name metadata?
has_gene_metadata <- function(x) {
  any(vapply(x, function(v) {
    gm <- attr(v, "gene_map")
    !is.null(gm) &&
      nrow(gm) > 0 &&
      "GeneName" %in% names(gm) &&
      any(!is.na(gm$GeneName) & gm$GeneName != "")
  }, logical(1)))
}

# Pull Gene Name values for CpGs.
# If set_name is provided, use only that set's metadata.
# If set_name is NULL, combine metadata from all selected sets.
get_gene_names_for_cpgs <- function(x, cpgs, set_name = NULL) {

  sets_to_use <- if (is.null(set_name)) names(x) else set_name

  meta_list <- list()

  for (s in sets_to_use) {
    gm <- attr(x[[s]], "gene_map")

    if (!is.null(gm) && nrow(gm) > 0) {
      gm$Set <- s
      meta_list[[s]] <- gm
    }
  }

  if (length(meta_list) == 0) {
    return(rep(NA_character_, length(cpgs)))
  }

  meta <- dplyr::bind_rows(meta_list)

  meta$GeneName[is.na(meta$GeneName)] <- ""

  meta_summary <- meta |>
    dplyr::filter(CpG %in% cpgs, GeneName != "") |>
    dplyr::group_by(CpG) |>
    dplyr::summarise(
      `Gene Name` = paste(unique(GeneName), collapse = "; "),
      .groups = "drop"
    )

  out <- rep(NA_character_, length(cpgs))

  idx <- match(cpgs, meta_summary$CpG)
  has_match <- !is.na(idx)

  out[has_match] <- meta_summary$`Gene Name`[idx[has_match]]

  out
}

# Export to Excel: Summary sheet + Set sheets + Intersect sheets
write_venn_excel <- function(x, vtable, file = "Venn_Sets_and_Intersections.xlsx") {
  wb   <- openxlsx::createWorkbook()
  bold <- openxlsx::createStyle(textDecoration = "bold")

  # Intersections (drop "(none)"), order by size desc then name
  intersections <- vtable |>
    dplyr::filter(intersect != "(none)") |>
    dplyr::arrange(dplyr::desc(n), intersect)

  # 🔧 sanity: ensure unique labels (guard against any upstream duplication)
  if (anyDuplicated(intersections$intersect)) {
    intersections <- intersections[!duplicated(intersections$intersect), ]
  }

  # Map: intersect -> genes (as character vector)
  inter_genes <- setNames(lapply(intersections$genes, unlist), intersections$intersect)

  # Map each gene -> its exclusive intersect (fast lookup)
  gene_to_intersect <- unlist(lapply(names(inter_genes), function(k) {
    setNames(rep.int(k, length(inter_genes[[k]])), inter_genes[[k]])
  }), use.names = TRUE)

  # Light pastel colors for up to 20 intersections
  pastel <- c(
    "#E9F2FF","#FFEDE6","#E6FFF2","#FFF8E6","#F1E6FF",
    "#E6FAFF","#FFF0F6","#F6FFE6","#E6FFFF","#FAE6FF",
    "#E6EEFF","#FFFDE6","#E6FFF9","#F3E6FF","#FFE6F3",
    "#EAF7FF","#FFF4E6","#E6FFF0","#F7E6FF","#F0FFF6"
  )
  if (nrow(intersections) > length(pastel)) {
    pastel <- rep(pastel, length.out = nrow(intersections))
  }
  color_map <- setNames(pastel[seq_len(nrow(intersections))], intersections$intersect)

  ## ---- SUMMARY SHEET -----------------------------------------------------

  # Define the expected "Set n Only" order from the current set names
  single_levels <- paste0(names(x), " Only")

  # Single-set-only intersections: "Set 1 Only", "Set 2 Only", ...
  singletons <- intersections |>
    dplyr::filter(intersect %in% single_levels) |>
    dplyr::mutate(.order = match(intersect, single_levels)) |>
    dplyr::arrange(.order) |>
    dplyr::select(Group = intersect, Count = n)

  # All other (multi-set) intersections
  multis <- intersections |>
    dplyr::filter(!intersect %in% single_levels) |>
    dplyr::select(Group = intersect, Count = n)

  # Bind rows: singles first in Set order, then all remaining
  summary_df <- dplyr::bind_rows(singletons, multis)

  # Add Summary sheet first by creating it before all others
  openxlsx::addWorksheet(wb, "Summary")
  openxlsx::writeData(wb, "Summary", summary_df, headerStyle = bold)

  ## ---- SHEET NAME TRACKING ----------------------------------------------

  # Keep track to avoid duplicate sheet names; start with "Summary"
  used_sheet_names <- "Summary"

  # Helper to make a safe/unique Excel sheet name
  make_sheet_name <- function(name) {
    nm <- gsub("[:\\\\/?*\\[\\]]", "-", name)  # forbidden chars
    if (nchar(nm) > 31) nm <- substr(nm, 1, 31)  # Excel limit
    base <- nm; k <- 1
    while (nm %in% used_sheet_names) {
      k <- k + 1
      suffix <- paste0("~", k)
      trim_len <- 31 - nchar(suffix)
      nm <- paste0(substr(base, 1, max(1, trim_len)), suffix)
    }
    used_sheet_names <<- c(used_sheet_names, nm)
    nm
  }

    ## ---- SET SHEETS --------------------------------------------------------

  include_gene_names <- has_gene_metadata(x)

  # For each Set sheet: columns = CpG + optional Gene Name + Intersect + all intersections that INCLUDE this Set
  for (s in names(x)) {
    genes_in_set <- sort_gene_asc(unique(unlist(x[s], use.names = FALSE)))

    # Intersections that contain this Set, including "Set n Only"
    inter_for_s <- intersections$intersect[
      vapply(intersections$intersect, function(lbl) {
        s %in% base_sets_in_label(lbl)
      }, logical(1))
    ]

    # Base output columns
    base_df <- data.frame(
      CpG = genes_in_set,
      check.names = FALSE
    )

    if (include_gene_names) {
      base_df[["Gene Name"]] <- get_gene_names_for_cpgs(
        x,
        genes_in_set,
        set_name = s
      )
    }

    base_df[["Intersect"]] <- unname(gene_to_intersect[genes_in_set])

    # Indicator matrix over those intersects
    if (length(inter_for_s) > 0) {
      ind_mat <- do.call(cbind, lapply(inter_for_s, function(inter_name) {
        as.integer(genes_in_set %in% inter_genes[[inter_name]])
      }))

      ind_df <- as.data.frame(ind_mat, check.names = FALSE)
      names(ind_df) <- inter_for_s

      df <- cbind(base_df, ind_df)
    } else {
      df <- base_df
    }

    # Write the sheet
    sh <- make_sheet_name(s)
    addWorksheet(wb, sh)
    writeData(wb, sh, df, headerStyle = bold)

    intersect_col <- which(names(df) == "Intersect")

    # Color the Intersect column cells by that intersect's color
    if (nrow(df) > 0 && length(intersect_col) == 1) {
      u_int <- unique(df$Intersect)

      for (int in u_int) {
        if (int %in% names(color_map)) {
          rows_i <- which(df$Intersect == int) + 1

          addStyle(
            wb, sh,
            style = createStyle(fgFill = color_map[[int]]),
            rows = rows_i,
            cols = intersect_col,
            gridExpand = TRUE,
            stack = TRUE
          )
        }
      }
    }

    # Color the intersect indicator columns: header + cells with 1
    if (length(inter_for_s) > 0) {
      for (cname in inter_for_s) {
        col_index <- match(cname, names(df))

        if (!is.na(col_index) && cname %in% names(color_map)) {
          addStyle(
            wb, sh,
            style = createStyle(textDecoration = "bold", fgFill = color_map[[cname]]),
            rows = 1,
            cols = col_index,
            gridExpand = TRUE,
            stack = TRUE
          )

          if (nrow(df) > 0) {
            rows_ones <- which(df[[cname]] == 1) + 1

            if (length(rows_ones)) {
              addStyle(
                wb, sh,
                style = createStyle(fgFill = color_map[[cname]]),
                rows = rows_ones,
                cols = col_index,
                gridExpand = TRUE,
                stack = TRUE
              )
            }
          }
        }
      }
    }
  }

    ## ---- INTERSECTION SHEETS ----------------------------------------------

  # One sheet per exclusive intersection; CpGs sorted
  for (i in seq_len(nrow(intersections))) {
    inter_label <- intersections$intersect[i]
    genes_vec   <- sort_gene_asc(unlist(intersections$genes[i]))

    df <- data.frame(
      CpG = genes_vec,
      check.names = FALSE
    )

    if (include_gene_names) {
      df[["Gene Name"]] <- get_gene_names_for_cpgs(
        x,
        genes_vec,
        set_name = NULL
      )
    }

    sh <- make_sheet_name(paste0("Intersect - ", inter_label))
    openxlsx::addWorksheet(wb, sh)
    openxlsx::writeData(wb, sh, df, headerStyle = bold)

    if (inter_label %in% names(color_map)) {
      openxlsx::addStyle(
        wb, sh,
        style = openxlsx::createStyle(
          textDecoration = "bold",
          fgFill = color_map[[inter_label]]
        ),
        rows = 1,
        cols = seq_len(ncol(df)),
        gridExpand = TRUE,
        stack = TRUE
      )
    }
  }
    openxlsx::saveWorkbook(wb, file, overwrite = TRUE)
  file
}
# Ensure tcltk is available for GUI
if (!requireNamespace("tcltk", quietly = TRUE)) {
  stop("The 'tcltk' package is required for the GUI comparison selector.")
}
library(tcltk)

# GUI: pick and queue multiple Venn comparisons with checkboxes
select_venn_comparisons_gui <- function(all_set_names, max_sets = 4) {
  if (!requireNamespace("tcltk", quietly = TRUE)) {
    stop("The 'tcltk' package is required for the GUI comparison selector.")
  }
  library(tcltk)

  tt <- tktoplevel()
  tkwm.title(tt, "Venn comparisons")

  queued <- list()

  # --- LEFT: checkboxes for all groups ---
  lbl_sets <- tklabel(tt, text = sprintf("Select 2–%d groups per comparison:", max_sets))

  frm_checks <- tkframe(tt)
  vars <- vector("list", length(all_set_names))

  for (i in seq_along(all_set_names)) {
    v <- tclVar("0")
    vars[[i]] <- v
    cb <- tkcheckbutton(
      frm_checks,
      text     = all_set_names[i],
      variable = v,
      onvalue  = "1",
      offvalue = "0"
    )
    tkpack(cb, anchor = "w")
  }

  # --- RIGHT: queued comparisons list ---
  lbl_queue <- tklabel(tt, text = "Queued comparisons:")
  lb_queue <- tklistbox(
    tt,
    selectmode      = "browse",
    height          = 10,
    exportselection = FALSE
  )

  # Variable to signal when Run or Cancel is pressed
  done_var <- tclVar(0)

  # Refresh the numbered list of queued comparisons
  refresh_queue <- function() {
    tkdelete(lb_queue, 0, "end")
    if (length(queued) == 0) return()
    for (i in seq_along(queued)) {
      line <- paste0(i, ") ", paste(queued[[i]], collapse = " ∩ "))
      tkinsert(lb_queue, "end", line)
    }
  }

  # --- Add selection button ---
  add_fun <- function() {
    # Which checkboxes are ON? (logical vapply, not integer)
    sel_idx <- which(vapply(
      vars,
      function(v) tclvalue(v) == "1",
      logical(1)
    ))

    if (length(sel_idx) < 2) {
      tkmessageBox(
        message = "Please select at least TWO groups for a Venn comparison.",
        icon    = "warning",
        type    = "ok"
      )
      return()
    }
    if (length(sel_idx) > max_sets) {
      tkmessageBox(
        message = sprintf("Please select AT MOST %d groups.", max_sets),
        icon    = "warning",
        type    = "ok"
      )
      return()
    }

    sel <- all_set_names[sel_idx]
    queued <<- c(queued, list(sel))
    refresh_queue()

    # Clear selections so you start fresh for the next comparison
    for (v in vars) tclvalue(v) <<- "0"
  }

  # --- Run button ---
  run_fun <- function() {
    if (length(queued) == 0) {
      tkmessageBox(
        message = "No comparisons queued.",
        icon    = "warning",
        type    = "ok"
      )
      return()
    }
    tclvalue(done_var) <<- 1
  }

  # --- Cancel button ---
  cancel_fun <- function() {
    queued <<- list()
    tclvalue(done_var) <<- 1
  }

  add_btn    <- tkbutton(tt, text = "Add selection",          command = add_fun)
  run_btn    <- tkbutton(tt, text = "Run queued comparisons", command = run_fun)
  cancel_btn <- tkbutton(tt, text = "Cancel",                 command = cancel_fun)

  # Layout with grid
  tkgrid(lbl_sets,  column = 0, row = 0, sticky = "w", padx = 5, pady = 5)
  tkgrid(lbl_queue, column = 1, row = 0, sticky = "w", padx = 5, pady = 5)

  tkgrid(frm_checks, column = 0, row = 1, rowspan = 3, padx = 5, pady = 5, sticky = "n")
  tkgrid(lb_queue,   column = 1, row = 1, rowspan = 3, padx = 5, pady = 5, sticky = "n")

  tkgrid(add_btn,    column = 0, row = 4, padx = 5, pady = 5)
  tkgrid(run_btn,    column = 1, row = 4, padx = 5, pady = 5, sticky = "w")
  tkgrid(cancel_btn, column = 1, row = 4, padx = 5, pady = 5, sticky = "e")

  # Block until Run or Cancel is pressed
  tkwait.variable(done_var)
  tkdestroy(tt)

  queued
}


# Helper: create Venn diagram PNG/PDF for a given x and filename prefix
create_venn_plot_files <- function(x, prefix = "Venn_Diagram") {
  base_fill <- unlist(cfg$colors)
  fill_vec  <- rep(base_fill, length.out = length(x))

  # ------------------------------------------------------------
  # Auto-size export canvas based on group/set name length
  # ------------------------------------------------------------
  set_names <- names(x)
  n_sets <- length(set_names)

  max_label_chars <- max(nchar(set_names), na.rm = TRUE)
  total_label_chars <- sum(nchar(set_names), na.rm = TRUE)

  # Start from original 8 x 8 inches
  plot_width  <- cfg$min_width
  plot_height <- cfg$min_height

  # Add extra width/height for long labels
  # Two-set Venns mostly need extra horizontal room.
  if (n_sets == 2) {
    plot_width  <- plot_width  + max(0, max_label_chars - 18) * 0.22
    plot_height <- plot_height + max(0, max_label_chars - 28) * 0.06
  }

  # Three- and four-set Venns can need both horizontal and vertical room.
  if (n_sets >= 3) {
    plot_width  <- plot_width  + max(0, max_label_chars - 16) * 0.14
    plot_height <- plot_height + max(0, max_label_chars - 20) * 0.10
  }

  # Add a smaller adjustment for lots of total label text
  plot_width <- plot_width + max(0, total_label_chars - 60) * 0.025

  # Clamp to reasonable values so exports do not become absurdly large
  plot_width  <- min(max(plot_width, cfg$min_width), cfg$max_width)
  plot_height <- min(max(plot_height, cfg$min_height), cfg$max_height)

  # Keep 300 DPI PNG output
  dpi <- cfg$dpi
  png_width_px  <- as.integer(round(plot_width * dpi))
  png_height_px <- as.integer(round(plot_height * dpi))

  # Slightly increase margin for long labels
  auto_margin <- if (max_label_chars > 25) 0.10 else 0.06
    # For rare long-name 2-set Venns where the external count label overlaps
  # a group label, move the external count label/connector to the bottom.
  auto_ext_text <- n_sets == 2 && max_label_chars >= 24
  auto_ext_pos  <- if (auto_ext_text) 180 else 0

  ## ---------- Labeled Venn (shown + PNG + PDF) ----------
  venn_grob <- VennDiagram::venn.diagram(
    x,
    filename       = NULL,   # return a grob so we can draw & save
    category.names = set_names,

    # Circles
    lwd            = 2,
    lty            = "blank",
    fill           = fill_vec,

    # Numbers (region counts)
    cex = cfg$count_size,
    fontface       = "bold",

    # Set names
    cat.cex = cfg$label_size,
    cat.fontface   = "bold",
    cat.default.pos= "outer",
        # Move external small-region count labels away from long group labels
    ext.text       = auto_ext_text,
    ext.pos        = auto_ext_pos,
    ext.line.lwd   = 2,
    ext.line.lty   = "solid",

    # Extra whitespace around the diagram to reduce clipping
    margin         = auto_margin
  )

  # Show labeled version in viewer
  # Interactive viewer disabled in the app.

  # Save PNG, auto-sized at 300 DPI
  png_path <- paste0(prefix, ".png")
  grDevices::png(
    filename = png_path,
    width    = png_width_px,
    height   = png_height_px,
    res      = dpi
  )
  grid::grid.draw(venn_grob)
  grDevices::dev.off()

  # Save labeled PDF, auto-sized
  pdf_path <- paste0(prefix, ".pdf")
  grDevices::pdf(
    file   = pdf_path,
    width  = plot_width,
    height = plot_height
  )
  grid::grid.draw(venn_grob)
  grDevices::dev.off()

  ## ---------- Unlabeled Venn (no group labels, numbers only) ----------
  venn_grob_unlabeled <- VennDiagram::venn.diagram(
    x,
    filename       = NULL,
    category.names = rep("", n_sets),

    # Circles
    lwd            = 2,
    lty            = "blank",
    fill           = fill_vec,

    # Numbers
    cex = cfg$count_size,
    fontface       = "bold",

    # Set names hidden
    cat.cex        = 0,
    cat.fontface   = "bold",
    cat.default.pos= "outer",
    cat.dist       = rep(0.08, n_sets),

    margin         = auto_margin
  )

  pdf_unlabeled_path <- paste0(prefix, "_unlabeled.pdf")
  grDevices::pdf(
    file   = pdf_unlabeled_path,
    width  = plot_width,
    height = plot_height
  )
  grid::grid.draw(venn_grob_unlabeled)
  grDevices::dev.off()

  cat(
    "Saved:",
    png_path, ",",
    pdf_path, "and",
    pdf_unlabeled_path,
    "\n"
  )

  cat(
    "Auto-sized export dimensions:",
    round(plot_width, 2), "x", round(plot_height, 2),
    "inches\n"
  )
}
# ===================== DRIVER: GUI queue and run all comparisons =====================

all_set_names <- names(set_list)

# This opens the popup window with:
# - left: list of groups, selectable (checkbox-style)
# - right: numbered list of queued comparisons
# - buttons: "Add selection", "Run queued comparisons", "Cancel"
comparison_list <- select_venn_comparisons_gui(all_set_names, max_sets = 4)

if (length(comparison_list) == 0) {
  cat("No comparisons to run.\n")
} else {
  
   # --- NEW: ensure output folders exist ---
  excel_dir <- "Sets and Intersections"
  plot_dir  <- "Venn Diagrams"

  if (!dir.exists(excel_dir)) dir.create(excel_dir, recursive = TRUE)
  if (!dir.exists(plot_dir))  dir.create(plot_dir,  recursive = TRUE)
  
  for (cmp_idx in seq_along(comparison_list)) {
    groups <- comparison_list[[cmp_idx]]

    cat("\nRunning comparison", cmp_idx, ":", paste(groups, collapse = " ∩ "), "\n")

    # Build the x list for this comparison
    x <- set_list[groups]
    names(x) <- groups  # ensure Venn labels match the file names

    # Venn table + Excel
    v.table <- venn_group_table(x)

    prefix_raw  <- paste(groups, collapse = "__")
    prefix_safe <- gsub("[^A-Za-z0-9_]+", "_", prefix_raw)

   # ----- Excel output (unchanged filename, now in "Sets and Intersections") -----
    excel_file <- paste0("Venn_", prefix_safe, ".xlsx")
    excel_path <- file.path(excel_dir, excel_file)
    xlsx_path  <- write_venn_excel(x, v.table, excel_path)
    cat("Excel written to:", xlsx_path, "\n")

    # ----- Venn plots (PNG + PDF, unchanged base name, now in "Venn Diagrams") -----
    plot_prefix <- file.path(plot_dir, paste0("Venn_", prefix_safe))
    create_venn_plot_files(x, prefix = plot_prefix)
  }
}
