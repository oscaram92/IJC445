# NUTS1 (GB)

# 0) Setup
pkgs <- c("readxl","dplyr","tidyr","stringr","readr","purrr","factoextra","ggplot2","cluster")
to_install <- pkgs[!vapply(pkgs, requireNamespace, FUN.VALUE = logical(1), quietly = TRUE)]
if (length(to_install)) install.packages(to_install)
invisible(lapply(pkgs, library, character.only = TRUE))

# improves repelled labels
has_ggrepel <- requireNamespace("ggrepel", quietly = TRUE)
if (has_ggrepel) suppressPackageStartupMessages(library(ggrepel))

set.seed(123)

# 1) Paths 
biz_file  <- "C:/Users/user/Documents/MSC Data Science Hub/IJC437 Introduction to Data Science/Report/businessdemographyexceltables2023.xlsx"
bres_file <- "C:/Users/user/Documents/MSC Data Science Hub/IJC437 Introduction to Data Science/Report/nomis_2025_12_04_201028.xlsx"


# 2) Constants 
years_panel <- 2018:2022
use_log_size <- TRUE   
k_opt <- 2             

clean_region <- function(x) stringr::str_trim(stringr::str_to_upper(as.character(x)))

nuts1_regions <- clean_region(c(
  "North East","North West","Yorkshire and The Humber","East Midlands","West Midlands",
  "East","London","South East","South West","Wales","Scotland"
))

repair_blank_names <- function(df) {
  n <- names(df)
  i <- which(is.na(n) | n == "")
  if (length(i)) names(df)[i] <- paste0("V", i)
  df
}

assert_complete_regions <- function(df, region_col = "region") {
  miss <- setdiff(nuts1_regions, df[[region_col]])
  if (length(miss)) stop("Missing regions after joins: ", paste(miss, collapse = ", "), call. = FALSE)
  invisible(TRUE)
}
assert_no_na <- function(df, msg = "NA values found") {
  if (anyNA(df)) stop(msg, call. = FALSE)
  invisible(TRUE)
}


# 3) ONS: multi-year (Births/Deaths/Actives)
read_ons_panel <- function(file, sheet, years = years_panel, region_col = 2) {
  year_min <- min(years); year_max <- max(years)
  
  df <- readxl::read_excel(file, sheet = sheet, col_names = FALSE) |>
    dplyr::filter(dplyr::if_any(dplyr::everything(), ~ !is.na(.)))
  
  num_mat <- suppressWarnings(as.data.frame(lapply(df, as.numeric)))
  has_year <- apply(num_mat, 1, function(r) any(dplyr::between(r, year_min, year_max), na.rm = TRUE))
  
  year_row <- which(is.na(df[[region_col]]) & has_year)[1]
  if (is.na(year_row)) stop("Year header row not found in sheet: ", sheet, call. = FALSE)
  
  years_row <- suppressWarnings(as.numeric(df[year_row, ]))
  year_cols <- which(dplyr::between(years_row, year_min, year_max))
  if (!length(year_cols)) stop("No year columns found in sheet: ", sheet, call. = FALSE)
  
  out <- df[(year_row + 1):nrow(df), c(region_col, year_cols)]
  names(out) <- c("region_raw", as.character(years_row[year_cols]))
  
  out |>
    dplyr::mutate(region = clean_region(region_raw)) |>
    dplyr::select(-region_raw) |>
    tidyr::pivot_longer(
      cols = -region, names_to = "year", values_to = "value",
      values_transform = list(value = as.character)
    ) |>
    dplyr::mutate(
      year  = as.integer(year),
      value = suppressWarnings(readr::parse_number(value))
    ) |>
    dplyr::filter(region %in% nuts1_regions, year %in% years, !is.na(value))
}

build_long_from_sheets <- function(file, sheets, years = years_panel, df_name = "series_long") {
  out <- purrr::map_dfr(sheets, ~ read_ons_panel(file, .x, years = years)) |>
    dplyr::mutate(year = as.integer(year))
  
    dup <- out |>
    dplyr::count(region, year, name = "n") |>
    dplyr::filter(n > 1)
  if (nrow(dup)) {
    stop("Duplicates detected in ", df_name, " (region-year):\n",
         paste(utils::capture.output(print(dup)), collapse = "\n"), call. = FALSE)
  }
  
    expected <- tidyr::expand_grid(region = nuts1_regions, year = years)
  missing <- expected |>
    dplyr::anti_join(out |> dplyr::distinct(region, year), by = c("region","year"))
  if (nrow(missing)) {
    stop("Missing region-year combinations in ", df_name, ":\n",
         paste(utils::capture.output(print(missing)), collapse = "\n"), call. = FALSE)
  }
  
  out
}


# 4) (ONS Table 4.1)
read_survival_3y <- function(file, sheet = "Table 4.1", skip = 4, survival_col = 8) {
  raw <- suppressMessages(readxl::read_excel(file, sheet = sheet, skip = skip, .name_repair = "minimal")) |>
    repair_blank_names()
  names(raw)[1] <- "region_raw"
  if (ncol(raw) < survival_col) stop("Sheet '", sheet, "' has fewer than ", survival_col, " columns.", call. = FALSE)
  
  surv_vec <- raw[[survival_col]]
  out <- raw |>
    dplyr::mutate(
      region = clean_region(region_raw),
      survival_3y_reg = suppressWarnings(readr::parse_number(as.character(surv_vec)))
    ) |>
    dplyr::filter(region %in% nuts1_regions, !is.na(survival_3y_reg)) |>
    dplyr::group_by(region) |>
    dplyr::slice(1) |>
    dplyr::ungroup() |>
    dplyr::select(region, survival_3y_reg)
  
  out$survival_3y_reg[out$survival_3y_reg > 1] <- out$survival_3y_reg[out$survival_3y_reg > 1] / 100
  out
}


# 5) (ONS Table 6.1)
read_employer_active_2023 <- function(file, sheet = "Table 6.1", skip = 4, region_col = 2) {
  raw <- suppressMessages(readxl::read_excel(file, sheet = sheet, skip = skip)) |>
    repair_blank_names()
  if (region_col > ncol(raw)) stop("region_col out of range in ", sheet, call. = FALSE)
  names(raw)[region_col] <- "region_raw"
  
  active_col <- which(tolower(trimws(names(raw))) == "active")[1]
  if (is.na(active_col)) stop("Column 'Active' not found in ", sheet, call. = FALSE)
  active_name <- names(raw)[active_col]
  
  raw |>
    dplyr::filter(!is.na(region_raw)) |>
    dplyr::mutate(
      region = clean_region(region_raw),
      employer_active_2023_reg = suppressWarnings(readr::parse_number(as.character(.data[[active_name]])))
    ) |>
    dplyr::filter(region %in% nuts1_regions, !is.na(employer_active_2023_reg)) |>
    dplyr::select(region, employer_active_2023_reg)
}

# 6) BRES employment 2018–2022 

read_bres_employment <- function(file, sheet = "Data", years = years_panel) {
  year_min <- min(years); year_max <- max(years)
  
  df <- readxl::read_excel(file, sheet = sheet, col_names = FALSE)
  names(df) <- paste0("V", seq_len(ncol(df)))
  
  date_rows <- which(tolower(as.character(df$V1)) == "date")
  if (!length(date_rows)) stop("No rows with 'date' found in BRES sheet.", call. = FALSE)
  
  n <- nrow(df); p <- ncol(df)
  res <- vector("list", 0)
  
  for (idx in seq_along(date_rows)) {
    start_row <- date_rows[idx]
    end_row <- if (idx < length(date_rows)) date_rows[idx + 1] - 1 else n
    
    year_val <- suppressWarnings(as.integer(df[start_row, 2, drop = TRUE]))
    if (is.na(year_val) || year_val < year_min || year_val > year_max) next
    
    rel <- which(as.character(df$V1[(start_row + 1):end_row]) == "Industry")[1]
    if (is.na(rel)) next
    header_row <- start_row + rel
    
    header_vals <- as.character(unlist(df[header_row, , drop = FALSE]))
    region_cols <- which(!is.na(header_vals) & header_vals != "Flags" & seq_len(p) > 1)
    if (!length(region_cols)) next
    
    sub <- df[(header_row + 2):end_row, , drop = FALSE]
    sub <- sub[!is.na(sub$V1) & as.character(sub$V1) != "NA", , drop = FALSE]
    
    industry_labels <- tolower(trimws(as.character(sub$V1)))
    total_idx <- which(industry_labels %in% c("total","all industries","all industries total","all"))
    
    for (j in region_cols) {
      reg <- clean_region(header_vals[j])
      if (!(reg %in% nuts1_regions)) next
      
      col_vals <- suppressWarnings(as.numeric(as.character(sub[[j]])))
      total_val <- if (length(total_idx) && !is.na(col_vals[total_idx[1]])) col_vals[total_idx[1]] else sum(col_vals, na.rm = TRUE)
      
      res[[length(res) + 1]] <- tibble::tibble(region = reg, year = year_val, value = total_val)
    }
  }
  
  if (!length(res)) stop("No BRES observations built; check sheet structure.", call. = FALSE)
  
  dplyr::bind_rows(res) |>
    dplyr::group_by(region, year) |>
    dplyr::summarise(value = sum(value, na.rm = TRUE), .groups = "drop")
}

# 7) Pooled rate (weighted average 2018–2022)

mean_rate_weighted <- function(num_long, den_long, years = years_panel, out_name = "rate_reg") {
  df <- num_long |>
    dplyr::inner_join(den_long, by = c("region","year"), suffix = c("_num","_den")) |>
    dplyr::filter(year %in% years, value_den > 0) |>
    dplyr::group_by(region) |>
    dplyr::summarise(
      n_years = dplyr::n_distinct(year),
      rate = sum(value_num, na.rm = TRUE) / sum(value_den, na.rm = TRUE),
      .groups = "drop"
    )
  
  if (any(df$n_years < length(years))) {
    bad <- df |> dplyr::filter(n_years < length(years))
    stop("Incomplete year coverage in mean_rate_weighted():\n",
         paste(utils::capture.output(print(bad)), collapse = "\n"), call. = FALSE)
  }
  
  df <- df |> dplyr::select(region, rate)
  names(df)[2] <- out_name
  df
}


# 8) RQ1 dataset 
birth_sheets  <- c("Table 1.1a","Table 1.1b","Table 1.1c","Table 1.1d")
death_sheets  <- c("Table 2.1a","Table 2.1b","Table 2.1c","Table 2.1d")
active_sheets <- c("Table 3.1a","Table 3.1b","Table 3.1c","Table 3.1d")

births_long <- build_long_from_sheets(biz_file, birth_sheets,  years_panel, "births_long")
deaths_long <- build_long_from_sheets(biz_file, death_sheets,  years_panel, "deaths_long")
active_long <- build_long_from_sheets(biz_file, active_sheets, years_panel, "active_long")

survival_df <- read_survival_3y(biz_file, "Table 4.1", skip = 4, survival_col = 8)
employer_active_2023 <- read_employer_active_2023(biz_file, "Table 6.1", skip = 4, region_col = 2)

emp_long <- read_bres_employment(bres_file, "Data", years_panel)

births_rates <- mean_rate_weighted(births_long, active_long, years_panel, "birth_rate_reg")
deaths_rates <- mean_rate_weighted(deaths_long, active_long, years_panel, "death_rate_reg")

emp_18_22 <- emp_long |>
  dplyr::filter(year %in% c(2018, 2022)) |>
  dplyr::select(region, year, value) |>
  tidyr::pivot_wider(names_from = year, values_from = value, names_prefix = "emp_") |>
  dplyr::filter(!is.na(emp_2018), !is.na(emp_2022), emp_2018 > 0) |>
  dplyr::mutate(emp_growth_18_22_reg = (emp_2022 - emp_2018) / emp_2018) |>
  dplyr::select(region, emp_growth_18_22_reg)

reg_cluster_df <- list(births_rates, deaths_rates, survival_df, emp_18_22, employer_active_2023) |>
  purrr::reduce(~ dplyr::inner_join(.x, .y, by = "region"))

assert_complete_regions(reg_cluster_df)
assert_no_na(reg_cluster_df, "NA values found in reg_cluster_df; decide on imputation or fix joins first.")


# 9) Clustering prep + K diagnostics

vars_for_clust <- reg_cluster_df |>
  dplyr::mutate(
    employer_active_2023_reg = if (use_log_size) log1p(employer_active_2023_reg) else employer_active_2023_reg
  ) |>
  dplyr::select(birth_rate_reg, death_rate_reg, survival_3y_reg, emp_growth_18_22_reg, employer_active_2023_reg)

reg_scaled <- scale(vars_for_clust)
row.names(reg_scaled) <- reg_cluster_df$region
dist_reg <- dist(reg_scaled, method = "euclidean")

sil_scores <- sapply(2:5, function(k) {
  km_tmp <- kmeans(reg_scaled, centers = k, nstart = 25)
  mean(cluster::silhouette(km_tmp$cluster, dist_reg)[, 3])
})
names(sil_scores) <- paste0("K=", 2:5)
print("Silhouette Scores (K-means):"); print(round(sil_scores, 3))

print(factoextra::fviz_nbclust(reg_scaled, kmeans, method = "wss", k.max = 5) +
        ggplot2::theme_minimal() + ggplot2::ggtitle("Elbow Method"))

print(factoextra::fviz_nbclust(reg_scaled, kmeans, method = "silhouette", k.max = 5) +
        ggplot2::theme_minimal() + ggplot2::ggtitle("Silhouette Method"))

gap_stat <- cluster::clusGap(reg_scaled, FUN = kmeans, nstart = 25, K.max = 5, B = 100)
print(factoextra::fviz_gap_stat(gap_stat) + ggplot2::theme_minimal() + ggplot2::ggtitle("Gap Statistic"))


# 10) k-means + HC +  visualizations (RQ1)
km <- kmeans(reg_scaled, centers = k_opt, nstart = 25)
hc_ward <- hclust(dist_reg, method = "ward.D2")
hc_clusters <- cutree(hc_ward, k = k_opt)

reg_cluster_df <- reg_cluster_df |>
  dplyr::mutate(
    cluster_kmeans = factor(km$cluster),
    cluster_hc     = factor(hc_clusters)
  )

print(factoextra::fviz_dend(
  hc_ward, k = k_opt, rect = TRUE, rect_fill = TRUE, color_labels_by_k = TRUE,
  horiz = TRUE, cex = 0.6, main = "Hierarchical Clustering Dendrogram",
  ggtheme = ggplot2::theme_minimal()
))

print(factoextra::fviz_cluster(
  list(data = reg_scaled, cluster = km$cluster),
  geom = "point", ellipse.type = "norm", ellipse.alpha = 0.15,
  show.clust.cent = TRUE, repel = TRUE, labelsize = 3,
  ggtheme = ggplot2::theme_minimal(),
  main = "K-means Clusters (PCA Space)"
))

sil_km <- cluster::silhouette(km$cluster, dist_reg)
print(factoextra::fviz_silhouette(sil_km) + ggplot2::theme_minimal() +
        ggplot2::ggtitle("Silhouette Plot (K-means)"))


# 11) RQ2 — descriptives by cluster
vars_orig <- c("birth_rate_reg","death_rate_reg","survival_3y_reg","emp_growth_18_22_reg","employer_active_2023_reg")

var_labels_en <- c(
  birth_rate_reg           = "birth rate",
  death_rate_reg           = "death rate",
  survival_3y_reg          = "3-year survival",
  emp_growth_18_22_reg     = "employment growth (2018–2022)",
  employer_active_2023_reg = "active employer enterprises (2023)"
)

cluster_summary <- reg_cluster_df |>
  dplyr::group_by(cluster_kmeans) |>
  dplyr::summarise(
    n_regions = dplyr::n(),
    dplyr::across(dplyr::all_of(vars_orig), ~ mean(.x, na.rm = TRUE), .names = "mean_{.col}"),
    .groups = "drop"
  )

overall_mean <- sapply(reg_cluster_df[vars_orig], mean, na.rm = TRUE)
overall_sd   <- sapply(reg_cluster_df[vars_orig], sd,   na.rm = TRUE)
overall_sd[overall_sd == 0] <- 1

cluster_summary <- cluster_summary |>
  dplyr::rowwise() |>
  dplyr::mutate(
    cluster_type = {
      means <- c_across(dplyr::starts_with("mean_"))
      names(means) <- sub("^mean_", "", names(means))
      z <- (means[vars_orig] - overall_mean) / overall_sd
      hi <- names(sort(z, decreasing = TRUE))[1:2]
      lo <- names(sort(z, decreasing = FALSE))[1:2]
      paste(
        c(paste0("High ", var_labels_en[hi[1]]),
          paste0("High ", var_labels_en[hi[2]]),
          paste0("Low ",  var_labels_en[lo[1]]),
          paste0("Low ",  var_labels_en[lo[2]])),
        collapse = "; "
      )
    }
  ) |>
  dplyr::ungroup()

reg_cluster_df <- reg_cluster_df |>
  dplyr::left_join(cluster_summary |> dplyr::select(cluster_kmeans, cluster_type), by = "cluster_kmeans")

print("Region Assignment (RQ1):")
print(reg_cluster_df |> dplyr::select(region, cluster_kmeans, cluster_type))

print("Average Characteristics by Cluster (RQ2):")
print(cluster_summary)

cluster_summary %>%
  dplyr::select(
    cluster_kmeans, n_regions,
    mean_emp_growth_18_22_reg,
    mean_employer_active_2023_reg
  ) %>%
  print()

# 13) IJC445 / Geospatial visualization

if (!requireNamespace("ggmap", quietly = TRUE)) install.packages("ggmap")
library(ggmap)

register_stadiamaps("87xxx74d-a111-449x-x22x-xx2234101f66")

centroids_df <- data.frame(
  region = c(
    "NORTH EAST","NORTH WEST","YORKSHIRE AND THE HUMBER","EAST MIDLANDS","WEST MIDLANDS",
    "EAST","LONDON","SOUTH EAST","SOUTH WEST","WALES","SCOTLAND"
  ),
  lat = c(54.97, 53.48, 53.80, 52.95, 52.49, 52.20, 51.51, 51.28, 51.00, 52.30, 56.50),
  lon = c(-1.62, -2.24, -1.55, -1.15, -1.89, 0.12, -0.13, -0.43, -3.00, -3.70, -4.00),
  stringsAsFactors = FALSE
)

reg_cluster_df <- reg_cluster_df %>%
  dplyr::left_join(centroids_df, by = "region")

if (anyNA(reg_cluster_df$lat) || anyNA(reg_cluster_df$lon)) {
  stop("Missing lat/lon after centroid join; check region names.", call. = FALSE)
}

base_map <- get_stadiamap(
  bbox = c(left = -8, bottom = 50, right = 2, top = 59),
  zoom = 6,
  maptype = "stamen_terrain",
  color = "bw"
)

# 14) Cluster Profile Visualization (Z-Score Bars)
cluster_map_plot <- ggmap(base_map) +
  ggplot2::geom_point(
    data = reg_cluster_df,
    ggplot2::aes(x = lon, y = lat, color = cluster_kmeans),
    size = 3
  ) +
  ggplot2::theme_minimal() +
  ggplot2::ggtitle("Spatial distribution of clusters in Great Britain") +
  ggplot2::labs(color = "Cluster")

print(cluster_map_plot)



labs <- c(
  birth_rate_reg="Birth Rate", death_rate_reg="Death Rate",
  survival_3y_reg="3-Year Survival Rate", emp_growth_18_22_reg="Employment Growth (2018–2022)",
  employer_active_2023_reg="Active Employer Enterprises (2023)"
)

df <- reg_cluster_df %>%
  select(cluster_kmeans, all_of(names(labs))) %>%
  pivot_longer(-cluster_kmeans, names_to="var", values_to="value") %>%
  group_by(var) %>%
  mutate(s=sd(value,na.rm=TRUE),
         z=(value-mean(value,na.rm=TRUE))/ifelse(is.na(s)|s==0,1,s)) %>%
  ungroup() %>%
  group_by(cluster_kmeans, var) %>%
  summarise(z_mean=mean(z,na.rm=TRUE), .groups="drop") %>%
  mutate(cluster=factor(paste("Cluster", cluster_kmeans)),
         var=factor(var, levels=names(labs), labels=labs),
         lab=sprintf("%.2f", z_mean), hjust=ifelse(z_mean>=0,-0.15,1.15))

pd <- position_dodge(.75)

ggplot(df, aes(var, z_mean, fill=cluster)) +
  geom_hline(yintercept=0) +
  geom_col(position=pd, width=.7) +
  geom_text(aes(label=lab, hjust=hjust), position=pd, size=3) +
  coord_flip(clip="off") +
  labs(x=NULL, y="Mean z-score", fill=NULL, title="Cluster Profiles") +
  theme_minimal() + theme(legend.position="bottom")



print(reg_cluster_df)

reg_cluster_df %>%
  group_by(cluster_kmeans) %>%
  summarise(
    n = n(),
    birth = mean(birth_rate_reg),
    death = mean(death_rate_reg),
    survival = mean(survival_3y_reg),
    emp_growth = mean(emp_growth_18_22_reg),
    active_2023 = mean(employer_active_2023_reg),
    .groups = "drop"
  )

reg_cluster_df %>%
  select(region, cluster_kmeans) %>%
  arrange(cluster_kmeans, region)





























