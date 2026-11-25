# DNA Methylation Manhattan Plot
# Dataset columns expected:
# chr | pos | adj.P.Val | CpG

# Install once if needed:
# install.packages(c("readxl", "dplyr", "ggplot2", "stringr", "ggrepel"))

library(readxl)
library(dplyr)
library(ggplot2)
library(stringr)
library(ggrepel)
# 1. Load workbook / detect sheets

file_path <- file.choose()
file_name <- tools::file_path_sans_ext(basename(file_path))
out_dir <- dirname(file_path)

sheet_names <- readxl::excel_sheets(file_path)

# Helper function to make file/sheet names safe for output filenames
clean_file_stub <- function(x) {
  x %>%
    stringr::str_replace_all("[\\\\/:*?\"<>|]", "_") %>%
    stringr::str_replace_all("\\s+", "_") %>%
    stringr::str_replace_all("_+", "_") %>%
    stringr::str_replace_all("^_|_$", "")
}
# 2. Manhattan plot function

make_manhattan_plot <- function(file_path, sheet_name, output_name, out_dir) {

  message("Running Manhattan plot for sheet: ", sheet_name)

  raw_df <- readxl::read_excel(
    file_path,
    sheet = sheet_name,
    .name_repair = "unique"
  )

  supplied <- unlist(cfg$columns)
  if (anyDuplicated(supplied)) stop("Column mappings must be distinct.")
  if (!all(supplied %in% names(raw_df))) stop("Missing columns in sheet ", sheet_name, ": ", paste(setdiff(supplied, names(raw_df)), collapse=", "))
  raw_df <- raw_df[, supplied, drop=FALSE]
  names(raw_df) <- c("chr", "pos", "adj.P.Val", "CpG")
  required_cols <- c("chr", "pos", "adj.P.Val", "CpG")
  missing_cols <- setdiff(required_cols, names(raw_df))

  if (length(missing_cols) > 0) {
    warning(
      "Skipping sheet '", sheet_name,
      "' because these required columns are missing: ",
      paste(missing_cols, collapse = ", ")
    )
    return(NULL)
  }

  # ------------------------------------------------------------
  # Clean and format dataset
  # ------------------------------------------------------------

  manhattan_df <- raw_df %>%
    select(chr, pos, adj.P.Val, CpG) %>%
    rename(
      chr_raw = chr,
      position = pos,
      p_adj = adj.P.Val,
      CpG_ID = CpG
    ) %>%
    mutate(
      chr_clean = toupper(str_replace(trimws(as.character(chr_raw)), regex("^chr", ignore_case=TRUE), "")),
      chr_num = suppressWarnings(as.integer(chr_clean)),
      chr_order = case_when(
        !is.na(chr_num) & chr_num >= 1 & chr_num <= cfg$autosomes ~ chr_num,
        chr_clean == "X" ~ as.integer(cfg$autosomes + 1),
        chr_clean == "Y" ~ as.integer(cfg$autosomes + 2),
        chr_clean %in% c("M", "MT") ~ as.integer(cfg$autosomes + 3),
        TRUE ~ NA_integer_
      ),
      chr_label = case_when(
        chr_order <= cfg$autosomes ~ as.character(chr_order),
        chr_order == cfg$autosomes + 1 ~ "X",
        chr_order == cfg$autosomes + 2 ~ "Y",
        chr_order == cfg$autosomes + 3 ~ "MT",
        TRUE ~ NA_character_
      ),
      position = as.numeric(position),
      p_adj = as.numeric(p_adj)
    ) %>%
    filter(
      !is.na(chr_order),
      !is.na(position),
      !is.na(p_adj),
      p_adj > 0, p_adj <= 1, is.finite(position), position >= 0
    ) %>%
    mutate(
      neg_log10_p = -log10(p_adj)
    )

  message("Rows retained: ", nrow(manhattan_df), "/", nrow(raw_df), "; excluded: ", nrow(raw_df)-nrow(manhattan_df))
  if (nrow(manhattan_df) == 0) {
    warning("Skipping sheet '", sheet_name, "' because no usable rows were found.")
    return(NULL)
  }

  # ------------------------------------------------------------
  # Create equal-width chromosome plotting positions
  # ------------------------------------------------------------

  chr_display_width <- cfg$chr_width
  chr_gap <- cfg$chr_gap

  chr_info_df <- manhattan_df %>%
    distinct(chr_order, chr_label) %>%
    arrange(chr_order) %>%
    mutate(
      chr_index = row_number(),
      chr_start = (chr_index - 1) * (chr_display_width + chr_gap),
      chr_end = chr_start + chr_display_width,
      chr_center = chr_start + chr_display_width / 2
    )

  plot_df <- manhattan_df %>%
    left_join(
      chr_info_df,
      by = c("chr_order", "chr_label")
    ) %>%
    group_by(chr_order, chr_label) %>%
    mutate(
      chr_min_pos = min(position, na.rm = TRUE),
      chr_max_pos = max(position, na.rm = TRUE),
      within_chr_position = if_else(
        chr_max_pos > chr_min_pos,
        (position - chr_min_pos) / (chr_max_pos - chr_min_pos),
        0.5
      ),
      plot_position = chr_start + within_chr_position * chr_display_width,
      chr_color_group = factor(chr_index %% 2)
    ) %>%
    ungroup()

  # ------------------------------------------------------------
  # Define significance threshold and labels
  # ------------------------------------------------------------

  # Because this dataset uses adjusted p-values, this line represents FDR < 0.05.
  fdr_cutoff <- cfg$cutoff
  fdr_line <- -log10(fdr_cutoff)

  # Label the top CpG hits
  top_n_labels <- as.integer(cfg$top_labels)

  label_df <- plot_df %>%
    arrange(p_adj) %>%
    slice_head(n = top_n_labels)

  # ------------------------------------------------------------
  # Generate Manhattan plot
  # ------------------------------------------------------------

  manhattan_plot <- ggplot(
    plot_df,
    aes(
      x = plot_position,
      y = neg_log10_p,
      color = chr_color_group
    )
  ) +
    geom_point(alpha = cfg$alpha, size = cfg$point_size) +
    geom_hline(
      yintercept = fdr_line,
      linetype = "dashed",
      linewidth = 0.8,
      color = cfg$threshold_color
    ) +
    geom_text_repel(
      data = label_df,
      aes(
        x = plot_position,
        y = neg_log10_p,
        label = CpG_ID
      ),
      inherit.aes = FALSE,
      size = cfg$label_size,
      max.overlaps = Inf, seed = 1,
      box.padding = 0.35,
      point.padding = 0.25
    ) +
    scale_x_continuous(
      breaks = chr_info_df$chr_center,
      labels = chr_info_df$chr_label,
      limits = c(
        min(chr_info_df$chr_start),
        max(chr_info_df$chr_end)
      ),
      expand = expansion(mult = c(0.01, 0.01))
    ) +
    scale_y_continuous(
      expand = expansion(mult = c(0, 0.08))
    ) +
    scale_color_manual(
      values = unlist(cfg$colors),
      guide = "none"
    ) +
    labs(
      title = cfg$title,
      subtitle = paste0(output_name, " | y-axis = -log10(adjusted p-value)"),
      x = cfg$x_label,
      y = cfg$y_label
    ) +
    theme_classic(base_size = cfg$font_size) +
    theme(
      plot.title = element_text(face = "bold", hjust = 0.5),
      plot.subtitle = element_text(hjust = 0.5),
      axis.text.x = element_text(
        size = ifelse(n_distinct(chr_info_df$chr_label) > 25, 9, 11),
        angle = ifelse(n_distinct(chr_info_df$chr_label) > 25, 45, 0),
        hjust = ifelse(n_distinct(chr_info_df$chr_label) > 25, 1, 0.5)
      ),
      axis.text.y = element_text(size = 11),
      axis.title = element_text(face = "bold"),
      panel.border = element_blank()
    )

  # No interactive graphics device in the app.

  # ------------------------------------------------------------
  # Adaptive figure sizing
  # ------------------------------------------------------------

  get_manhattan_plot_size <- function(
      plot_df,
      chr_info_df,
      label_df = NULL,
      min_width = cfg$min_width,
      max_width = cfg$max_width,
      min_height = cfg$min_height,
      max_height = cfg$max_height
  ) {

    n_chr <- n_distinct(chr_info_df$chr_label)
    n_points <- nrow(plot_df)
    max_y <- max(plot_df$neg_log10_p, na.rm = TRUE)

    n_labels <- if (!is.null(label_df)) nrow(label_df) else 0
    longest_chr_label <- max(nchar(as.character(chr_info_df$chr_label)), na.rm = TRUE)

    # Width logic:
    # More chromosomes and longer chromosome labels need more horizontal space.
    # More total points slightly increases width, but not linearly.
    width <- min_width +
      (n_chr * 0.35) +
      (longest_chr_label * 0.15) +
      (log10(max(n_points, 10)) * 0.75)

    # Height logic:
    # Taller plots help when -log10(p) values are large or many labels are drawn.
    height <- min_height +
      (max_y * 0.08) +
      (n_labels * 0.08)

    width <- min(max(width, min_width), max_width)
    height <- min(max(height, min_height), max_height)

    return(
      list(
        width = width,
        height = height
      )
    )
  }

  plot_size <- get_manhattan_plot_size(
    plot_df = plot_df,
    chr_info_df = chr_info_df,
    label_df = label_df
  )

  plot_size

  # ------------------------------------------------------------
  # Save plot with adaptive dimensions
  # ------------------------------------------------------------

  safe_output_name <- clean_file_stub(output_name)

  ggsave(
    filename = file.path(out_dir, paste0(safe_output_name, paste0("_", cfg$suffix, ".tiff"))),
    plot = manhattan_plot,
    width = plot_size$width,
    height = plot_size$height,
    units = "in",
    dpi = cfg$dpi,
    compression = "lzw"
  )

  ggsave(
    filename = file.path(out_dir, paste0(safe_output_name, paste0("_", cfg$suffix, ".png"))),
    plot = manhattan_plot,
    width = plot_size$width,
    height = plot_size$height,
    units = "in",
    dpi = cfg$dpi
  )

  return(manhattan_plot)
}
# 3. Run single-sheet or multi-sheet workflow

if (length(sheet_names) == 1) {

  # Single-sheet workbook:
  # Use the original Excel file name for output naming.
  make_manhattan_plot(
    file_path = file_path,
    sheet_name = sheet_names[1],
    output_name = file_name,
    out_dir = out_dir
  )

} else {

  # Multi-sheet workbook:
  # Each sheet becomes its own dataset and output file.
  # The sheet name becomes the output file name.
  for (sheet_name in sheet_names) {

    make_manhattan_plot(
      file_path = file_path,
      sheet_name = sheet_name,
      output_name = sheet_name,
      out_dir = out_dir
    )
  }
}