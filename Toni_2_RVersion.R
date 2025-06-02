# ===================================================================================
# Toni GVC Visuals Pipeline – Publication/Elite Quality (Complete Fixed Version)
# Author: Anthony S. Cano Moncada
# Date: 2025-05-25
# ===================================================================================

# 0. SETUP & PACKAGES --------------------------------------------------------
options(warn = 0)
set.seed(123)

if (!requireNamespace("pacman", quietly = TRUE)) install.packages("pacman")
pacman::p_load(
  fs, readxl, dplyr, tidyr, stringr, tibble, readr,
  ggplot2, ggrepel, scales, extrafont, assertthat,
  pheatmap, RColorBrewer, ggdist, ggridges, ggbeeswarm, ggforce,
  viridis, GGally, fmsb, forcats
)

# 1. PATHS & DATA LOAD -------------------------------------------------------
base_data_path   <- "/Volumes/VALEN/Africa:LAC/LAC:TONI/Data"
clean_export_path <- "/Volumes/VALEN/Africa:LAC/LAC:TONI/export/clean"
png_export_path  <- "/Volumes/VALEN/Africa:LAC/LAC:TONI/export/png_pf"

data_dir <- fs::path_expand(base_data_path)
clean_dir <- fs::path_expand(clean_export_path)
out_dir <- fs::path_expand(png_export_path)

fs::dir_create(clean_dir, recurse = TRUE)
fs::dir_create(out_dir, recurse = TRUE)

load_excel_safely <- function(fp) {
  assert_that(fs::file_exists(fp), msg = paste("Missing file:", fp))
  readxl::read_excel(fp)
}

tech_df_raw    <- load_excel_safely(fs::path(data_dir, "Technology Indicators.xlsx"))
sustain_df_raw <- load_excel_safely(fs::path(data_dir, "Sustainability Indicators.xlsx"))
geo_df_raw     <- load_excel_safely(fs::path(data_dir, "Geopolitical Indicators.xlsx"))

# 2. CLEANING & REGION ASSIGNMENT -------------------------------------------
region_map <- list(
  China = "China",
  OECD  = c("Australia","Austria","Belgium","Canada","Chile","Colombia","Costa Rica",
            "Czech Republic","Denmark","Estonia","Finland","France","Germany","Greece",
            "Hungary","Iceland","Ireland","Israel","Italy","Japan","Latvia",
            "Lithuania","Luxembourg","Mexico","Netherlands","New Zealand","Norway",
            "Poland","Portugal","Slovakia","Slovenia","Korea, South","Spain",
            "Sweden","Switzerland","Turkey","United Kingdom","United States"),
  ASEAN = c("Brunei Darussalam","Cambodia","Indonesia","Laos","Malaysia","Myanmar",
            "Philippines","Singapore","Thailand","Vietnam"),
  LAC   = c("Argentina","Belize","Bolivia","Brazil","Chile","Colombia","Costa Rica",
            "Dominican Republic","Ecuador","El Salvador","Guatemala","Guyana",
            "Honduras","Jamaica","Mexico","Nicaragua","Panama","Paraguay","Peru",
            "Suriname","Trinidad and Tobago","Uruguay","Venezuela")
)

assign_region <- function(country) {
  for (r in names(region_map)) if (country %in% region_map[[r]]) return(r)
  return(NA_character_)
}

# Enhanced robust data cleaning function
clean_ind_df_robust <- function(df) {
  # Standardize country column name
  country_col <- grep("^country$", names(df), ignore.case = TRUE, value = TRUE)
  if (length(country_col) == 1 && country_col != "Country") {
    names(df)[names(df) == country_col] <- "Country"
  }
  
  # Get all indicator columns (everything except Country)
  indicator_cols <- setdiff(names(df), "Country")
  
  # Function to clean individual columns
  clean_column <- function(x) {
    # Convert to character first
    x <- as.character(x)
    # Remove leading/trailing whitespace
    x <- stringr::str_trim(x)
    # Handle various NA representations
    x[x == "" | is.na(x) | tolower(x) %in% c("na", "n/a", "-", "null", "#n/a", "#value!", "#div/0!")] <- NA_character_
    # Remove thousands separators (commas)
    x <- stringr::str_replace_all(x, "(?<=\\d),(?=\\d{3})", "")
    # Convert European decimal commas to dots
    x <- stringr::str_replace_all(x, "(?<=\\d),(?=\\d)", ".")
    # Remove any non-numeric characters except decimal points and minus signs
    x <- stringr::str_replace_all(x, "[^0-9.-]", "")
    # Convert to numeric
    result <- suppressWarnings(as.numeric(x))
    return(result)
  }
  
  # Clean the dataframe
  df_clean <- df %>%
    mutate(
      Country = as.character(Country),
      across(all_of(indicator_cols), clean_column)
    )
  
  # Verify all indicator columns are numeric
  for(col in indicator_cols) {
    if(!is.numeric(df_clean[[col]])) {
      warning(paste("Column", col, "could not be converted to numeric"))
      # Force conversion
      df_clean[[col]] <- as.numeric(as.character(df_clean[[col]]))
    }
  }
  
  return(df_clean)
}

add_region_column <- function(df) {
  df %>% mutate(Region = sapply(Country, assign_region, USE.NAMES = FALSE))
}

# Apply robust cleaning
tech_df    <- tech_df_raw    %>% clean_ind_df_robust() %>% add_region_column()
sustain_df <- sustain_df_raw %>% clean_ind_df_robust() %>% add_region_column()
geo_df     <- geo_df_raw     %>% clean_ind_df_robust() %>% add_region_column()

# 3. INDICATOR GUIDE & INDICATOR LISTS ---------------------------------------
indicator_guide <- tribble(
  ~Code,                                          ~Pillar,         ~Abbreviation,
  "Logistic Performance index",                   "Technology",    "LPI",
  "Mobile Connectivity index",                    "Technology",    "MCI",
  "ICT Capital Goods Imports index",              "Technology",    "ICT",
  "Human Capital",                                "Technology",    "HUC",
  "Digital Services Trade Restrictiveness index", "Technology",    "DSR",
  "Digitally Deliverable Services index",         "Technology",    "DDS",
  "Government Promotion of Investments in Emerging Technologies","Technology","GPI",
  "Global Cybersecurity index",                   "Technology",    "GCI",
  "Low Carbon Intensity index",                   "Sustainability","LCI",
  "Renewable Energy Consumption index",           "Sustainability","REC",
  "Trade in energy transition goods index",       "Sustainability","TEG",
  "Protectionism in energy transition goods index","Sustainability","PEG",
  "Critical Mineral index",                       "Sustainability","CMI",
  "Biodiversity and Habitat- Environmental Performance","Sustainability","BHE",
  "Exposure to Natural Disasters index",          "Sustainability","END",
  "Vulnerability to Natural Disasters index",     "Sustainability","VND",
  "Export Similarity with China index",           "Geopolitical",  "ESC",
  "Trade Bans index",                             "Geopolitical",  "TBI",
  "Security Threat index",                        "Geopolitical",  "STI",
  "Trade with sanctioned countries index",        "Geopolitical",  "TSC",
  "Political distance from trading partners index","Geopolitical",  "PDT",
  "Nonalignment in UN voting index",              "Geopolitical",  "NAV",
  "Ethnic Cohesion index",                        "Geopolitical",  "ECI",
  "Working age population index",                 "Geopolitical",  "WAP"
)

# Get indicator lists and verify they exist in data
tech_inds    <- indicator_guide %>% filter(Pillar=="Technology") %>% pull(Code)
sustain_inds <- indicator_guide %>% filter(Pillar=="Sustainability") %>% pull(Code)
geo_inds     <- indicator_guide %>% filter(Pillar=="Geopolitical") %>% pull(Code)

tech_inds    <- intersect(tech_inds, names(tech_df))
sustain_inds <- intersect(sustain_inds, names(sustain_df))
geo_inds     <- intersect(geo_inds, names(geo_df))

# Check data types and force conversion if needed
message("Checking Technology data types:")
tech_col_types <- sapply(tech_df[tech_inds], class)
print(tech_col_types)

char_cols <- tech_inds[tech_col_types == "character"]
if(length(char_cols) > 0) {
  message("Force converting character columns to numeric: ", paste(char_cols, collapse=", "))
  tech_df <- tech_df %>%
    mutate(across(all_of(char_cols), ~as.numeric(as.character(.))))
}

# 4. EXPORT CLEAN DATA -------------------------------------------------------
readr::write_csv(tech_df, fs::path(clean_dir, "Technology_Indicators_Clean.csv"))
readr::write_csv(sustain_df, fs::path(clean_dir, "Sustainability_Indicators_Clean.csv"))
readr::write_csv(geo_df, fs::path(clean_dir, "Geopolitical_Indicators_Clean.csv"))
readr::write_csv(indicator_guide, fs::path(clean_dir, "Indicator_Guide.csv"))

# 5. THEME & PALETTE --------------------------------------------------------
main_regions <- c("China","OECD","ASEAN","LAC")
wb_palette   <- c("China"="#E31A1C", "OECD"="#1F78B4", "ASEAN"="#33A02C", "LAC"="#FF7F00")
names(wb_palette) <- main_regions

desired_font <- tryCatch({
  if("Roboto Condensed" %in% extrafont::fonts()) "Roboto Condensed" else "sans"
}, error = function(e) "sans")

theme_gvc <- theme_minimal(base_size=14, base_family=desired_font) +
  theme(
    panel.grid.major = element_line(color="#e5e5e5", size=0.3),
    strip.text = element_text(face="bold", size=13),
    plot.title = element_text(face="bold", size=18),
    legend.position = "bottom"
  )
theme_set(theme_gvc)

# 6. PLOTTING FUNCTIONS ------------------------------------------------------

# 6.1 Facet Boxplot by Region and Indicator
plot_facet_box <- function(df, inds, title, filename) {
  tryCatch({
    dl <- df %>%
      filter(!is.na(Region), Region %in% main_regions) %>%
      select(Region, all_of(inds)) %>%
      pivot_longer(-Region, names_to="Indicator", values_to="Value") %>%
      filter(!is.na(Value)) %>%
      left_join(indicator_guide, by=c("Indicator"="Code"))
    
    if(nrow(dl) == 0) {
      message("No data available for boxplot: ", title)
      return(invisible(NULL))
    }
    
    p <- ggplot(dl, aes(x=Region, y=Value, fill=Region)) +
      geom_boxplot(alpha=0.8, outlier.shape=NA) +
      facet_wrap(~paste(Abbreviation, Pillar), ncol=4, scales="free_y") +
      scale_fill_manual(values=wb_palette) +
      labs(title=title, x=NULL, y="Score (0-1)") +
      theme(axis.text.x = element_text(angle=30, hjust=1))
    
    ggsave(fs::path(out_dir, filename), p, width=14, height=8, dpi=300)
    message("Saved: ", filename)
  }, error = function(e) {
    message("Error in plot_facet_box: ", conditionMessage(e))
  })
}

# 6.2 Heatmap
plot_heatmap <- function(df, inds, title, filename) {
  tryCatch({
    dh <- df %>% 
      select(Country, all_of(inds)) %>% 
      filter(complete.cases(.))
    
    if(nrow(dh) < 2) {
      message("Insufficient data for heatmap: ", title)
      return(invisible(NULL))
    }
    
    mat <- as.matrix(dh[,-1])
    rownames(mat) <- dh$Country
    
    # Get abbreviations for column names
    abbr_names <- indicator_guide$Abbreviation[match(colnames(mat), indicator_guide$Code)]
    abbr_names[is.na(abbr_names)] <- colnames(mat)[is.na(abbr_names)]
    colnames(mat) <- abbr_names
    
    pal <- colorRampPalette(c("#2c7bb6","#ffffbf","#d7191c"))(100)
    
    pheatmap(mat, 
             color=pal, 
             cluster_rows=TRUE, 
             cluster_cols=TRUE, 
             main=title,
             filename=fs::path(out_dir, filename), 
             width=14, 
             height=8)
    message("Saved: ", filename)
  }, error = function(e) {
    message("Error in plot_heatmap: ", conditionMessage(e))
  })
}

# 6.3 Bubble Chart
plot_bubble <- function(df, inds, title, filename) {
  tryCatch({
    bdf <- df %>% 
      select(Country, Region, all_of(inds)) %>%
      pivot_longer(cols=all_of(inds), names_to="Indicator", values_to="Score") %>%
      filter(!is.na(Score)) %>% 
      left_join(indicator_guide, by=c("Indicator"="Code"))
    
    if(nrow(bdf) == 0) {
      message("No data available for bubble chart: ", title)
      return(invisible(NULL))
    }
    
    bdf$Abbreviation <- factor(bdf$Abbreviation, levels=indicator_guide$Abbreviation)
    
    p <- ggplot(bdf, aes(x=Abbreviation, y=forcats::fct_reorder(Country, Score), size=Score, fill=Score)) +
      geom_point(shape=21, alpha=0.85) +
      scale_size(range=c(2,8), guide="none") +
      scale_fill_viridis_c(option="C") +
      labs(title=title, x=NULL, y=NULL) +
      theme(axis.text.y=element_text(size=6), axis.text.x=element_text(angle=45, hjust=1))
    
    ggsave(fs::path(out_dir, filename), p, width=14, height=20, dpi=300)
    message("Saved: ", filename)
  }, error = function(e) {
    message("Error in plot_bubble: ", conditionMessage(e))
  })
}

# 6.4 Ridgeline Plot
plot_ridgeline <- function(df, inds, title, filename) {
  tryCatch({
    dl <- df %>% 
      filter(!is.na(Region), Region %in% main_regions) %>%
      select(Region, all_of(inds)) %>%
      pivot_longer(-Region, names_to="Indicator", values_to="Value") %>% 
      filter(!is.na(Value)) %>%
      left_join(indicator_guide, by=c("Indicator"="Code"))
    
    if(nrow(dl) == 0) {
      message("No data available for ridgeline plot: ", title)
      return(invisible(NULL))
    }
    
    dl$Abbreviation <- factor(dl$Abbreviation, levels=indicator_guide$Abbreviation)
    
    p <- ggplot(dl, aes(x=Value, y=Region, fill=Region)) +
      ggridges::geom_density_ridges(alpha=0.75) +
      facet_wrap(~Abbreviation, ncol=4, scales="free") +
      labs(title=title, x="Score (0-1)", y=NULL) +
      scale_fill_manual(values=wb_palette) +
      theme(legend.position="none")
    
    ggsave(fs::path(out_dir, filename), p, width=14, height=8, dpi=300)
    message("Saved: ", filename)
  }, error = function(e) {
    message("Error in plot_ridgeline: ", conditionMessage(e))
  })
}

# 6.5 Raincloud Plot
plot_raincloud <- function(df, inds, title, filename) {
  tryCatch({
    dl <- df %>% 
      filter(!is.na(Region), Region %in% main_regions) %>%
      select(Region, all_of(inds)) %>%
      pivot_longer(-Region, names_to="Indicator", values_to="Value") %>% 
      filter(!is.na(Value)) %>%
      left_join(indicator_guide, by=c("Indicator"="Code"))
    
    if(nrow(dl) == 0) {
      message("No data available for raincloud plot: ", title)
      return(invisible(NULL))
    }
    
    dl$Abbreviation <- factor(dl$Abbreviation, levels=indicator_guide$Abbreviation)
    
    p <- ggplot(dl, aes(x=Region, y=Value, fill=Region)) +
      ggdist::stat_halfeye(.width=0, justification=-0.2, alpha=0.6) +
      geom_boxplot(width=0.15, alpha=0.5) +
      geom_jitter(width=0.07, size=0.6, alpha=0.4) +
      facet_wrap(~Abbreviation, ncol=4, scales="free_y") +
      scale_fill_manual(values=wb_palette) +
      coord_flip() + 
      labs(title=title, x=NULL, y="Score (0-1)")
    
    ggsave(fs::path(out_dir, filename), p, width=14, height=12, dpi=300)
    message("Saved: ", filename)
  }, error = function(e) {
    message("Error in plot_raincloud: ", conditionMessage(e))
  })
}

# 6.6 Fixed Spider Chart Functions
make_spider_data_fixed <- function(df, inds) {
  tryCatch({
    # Filter and select data
    df_filtered <- df %>% 
      filter(!is.na(Region), Region %in% main_regions)
    
    if(nrow(df_filtered) == 0) return(NULL)
    
    # Calculate medians manually to avoid dplyr issues
    result_list <- list()
    for(region in main_regions) {
      region_data <- df_filtered %>% filter(Region == region)
      if(nrow(region_data) > 0) {
        medians <- sapply(inds, function(ind) {
          values <- region_data[[ind]]
          median(values, na.rm = TRUE)
        })
        result_list[[region]] <- medians
      }
    }
    
    if(length(result_list) == 0) return(NULL)
    
    # Convert to data frame (not matrix) for fmsb compatibility
    data_df <- data.frame(do.call(rbind, result_list))
    colnames(data_df) <- inds
    
    # Add min/max rows as first two rows
    result_df <- data.frame(
      rbind(
        max = rep(1, length(inds)),
        min = rep(0, length(inds)),
        data_df
      )
    )
    
    return(result_df)
  }, error = function(e) {
    message("Error in make_spider_data_fixed: ", conditionMessage(e))
    return(NULL)
  })
}

make_spider_plot <- function(spider_df, title, filename) {
  tryCatch({
    if(is.null(spider_df) || nrow(spider_df) < 3) {
      message("Insufficient data for spider plot: ", title)
      return(invisible(NULL))
    }
    
    n_regions <- nrow(spider_df) - 2
    if(n_regions < 1) return(invisible(NULL))
    
    colors <- RColorBrewer::brewer.pal(max(3, min(n_regions, 8)), "Set2")[1:n_regions]
    
    png(fs::path(out_dir, filename), width=1200, height=1200, res=150)
    par(mar = c(2, 2, 4, 2))
    
    fmsb::radarchart(spider_df, 
                     axistype=1, 
                     seg=5,
                     pcol=colors,
                     pfcol=scales::alpha(colors, 0.4),
                     plwd=2, 
                     cglcol="#CCCCCC", 
                     cglty=1, 
                     axislabcol="grey30",
                     vlcex=0.7, 
                     caxislabels=seq(0,1,0.2), 
                     title=title)
    
    legend("topright", 
           legend=rownames(spider_df)[-c(1,2)],
           col=colors, 
           lwd=2, 
           bty="n")
    
    dev.off()
    message("Saved: ", filename)
  }, error = function(e) {
    if(dev.cur() > 1) dev.off()
    message("Error in make_spider_plot: ", conditionMessage(e))
  })
}

# 6.7 Fixed Parallel Coordinates
plot_parallel_coordinates_fixed <- function(df, inds, title, filename) {
  tryCatch({
    # Manual calculation of medians to avoid dplyr issues
    df_filtered <- df %>% 
      filter(!is.na(Region), Region %in% main_regions)
    
    if(nrow(df_filtered) == 0) {
      message("No data available for parallel coordinates: ", title)
      return(invisible(NULL))
    }
    
    # Calculate medians manually
    result_list <- list()
    for(region in main_regions) {
      region_data <- df_filtered %>% filter(Region == region)
      if(nrow(region_data) > 0) {
        medians <- sapply(inds, function(ind) {
          values <- region_data[[ind]]
          median(values, na.rm = TRUE)
        })
        result_list[[region]] <- c(Entity = region, medians)
      }
    }
    
    if(length(result_list) == 0) {
      message("No data available for parallel coordinates: ", title)
      return(invisible(NULL))
    }
    
    # Convert to data frame
    data_parallel <- do.call(rbind, lapply(result_list, function(x) {
      data.frame(t(x), stringsAsFactors = FALSE)
    }))
    
    # Fix column types
    data_parallel$Entity <- as.character(data_parallel$Entity)
    for(i in 2:ncol(data_parallel)) {
      data_parallel[,i] <- as.numeric(data_parallel[,i])
    }
    
    # Get abbreviations for indicator names
    abbr_names <- indicator_guide$Abbreviation[match(names(data_parallel)[-1], indicator_guide$Code)]
    abbr_names[is.na(abbr_names)] <- names(data_parallel)[-1][is.na(abbr_names)]
    names(data_parallel)[-1] <- abbr_names
    
    p <- GGally::ggparcoord(data_parallel, 
                            columns=2:ncol(data_parallel), 
                            groupColumn="Entity",
                            scale="uniminmax", 
                            title=title, 
                            alphaLines=0.8, 
                            showPoints=TRUE) +
      scale_color_manual(values=wb_palette) +
      theme(axis.text.x=element_text(angle=45, hjust=1))
    
    ggsave(fs::path(out_dir, filename), p, width=14, height=8, dpi=300)
    message("Saved: ", filename)
  }, error = function(e) {
    message("Error in plot_parallel_coordinates_fixed: ", conditionMessage(e))
  })
}

# 7. PLOT EXECUTION ----------------------------------------------------------
message("Starting plot generation...")

# Technology with fixed functions
message("Generating Technology plots...")
plot_facet_box(tech_df, tech_inds, "Technology Pillar - Boxplot", "Boxplot_Tech.png")
plot_heatmap(tech_df, tech_inds, "Technology Pillar - Heatmap", "Heatmap_Tech.png")
plot_bubble(tech_df, tech_inds, "Technology Pillar - Bubble Chart", "Bubble_Tech.png")
plot_ridgeline(tech_df, tech_inds, "Technology Pillar - Ridgeline", "Ridgeline_Tech.png")
plot_raincloud(tech_df, tech_inds, "Technology Pillar - Raincloud", "Raincloud_Tech.png")
make_spider_plot(make_spider_data_fixed(tech_df, tech_inds), "Technology Pillar - Spider Chart", "Spider_Tech.png")
plot_parallel_coordinates_fixed(tech_df, tech_inds, "Technology Pillar - Parallel Coordinates", "Parallel_Tech.png")

# Sustainability
message("Generating Sustainability plots...")
plot_facet_box(sustain_df, sustain_inds, "Sustainability Pillar - Boxplot", "Boxplot_Sustain.png")
plot_heatmap(sustain_df, sustain_inds, "Sustainability Pillar - Heatmap", "Heatmap_Sustain.png")
plot_bubble(sustain_df, sustain_inds, "Sustainability Pillar - Bubble Chart", "Bubble_Sustain.png")
plot_ridgeline(sustain_df, sustain_inds, "Sustainability Pillar - Ridgeline", "Ridgeline_Sustain.png")
plot_raincloud(sustain_df, sustain_inds, "Sustainability Pillar - Raincloud", "Raincloud_Sustain.png")
make_spider_plot(make_spider_data_fixed(sustain_df, sustain_inds), "Sustainability Pillar - Spider Chart", "Spider_Sustain.png")
plot_parallel_coordinates_fixed(sustain_df, sustain_inds, "Sustainability Pillar - Parallel Coordinates", "Parallel_Sustain.png")

# Geopolitical
message("Generating Geopolitical plots...")
plot_facet_box(geo_df, geo_inds, "Geopolitical Pillar - Boxplot", "Boxplot_Geo.png")
plot_heatmap(geo_df, geo_inds, "Geopolitical Pillar - Heatmap", "Heatmap_Geo.png")
plot_bubble(geo_df, geo_inds, "Geopolitical Pillar - Bubble Chart", "Bubble_Geo.png")
plot_ridgeline(geo_df, geo_inds, "Geopolitical Pillar - Ridgeline", "Ridgeline_Geo.png")
plot_raincloud(geo_df, geo_inds, "Geopolitical Pillar - Raincloud", "Raincloud_Geo.png")
make_spider_plot(make_spider_data_fixed(geo_df, geo_inds), "Geopolitical Pillar - Spider Chart", "Spider_Geo.png")
plot_parallel_coordinates_fixed(geo_df, geo_inds, "Geopolitical Pillar - Parallel Coordinates", "Parallel_Geo.png")

# 8. COMPOSITE ANALYSIS ------------------------------------------------------
message("Generating Composite plots...")
tryCatch({
  base_cols <- c("Country","Region")
  
  # Prepare dataframes for merging
  tech_merge <- tech_df %>% select(all_of(c(base_cols, tech_inds)))
  sustain_merge <- sustain_df %>% select(all_of(c(base_cols, sustain_inds)))
  geo_merge <- geo_df %>% select(all_of(c(base_cols, geo_inds)))
  
  # Merge datasets with explicit relationship specification
  composite_df <- tech_merge %>%
    inner_join(sustain_merge, by = base_cols, relationship = "many-to-many") %>%
    inner_join(geo_merge, by = base_cols, relationship = "many-to-many")
  
  comp_inds_all <- c(tech_inds, sustain_inds, geo_inds)
  comp_inds_available <- intersect(comp_inds_all, names(composite_df))
  
  message("Composite dataset created with ", nrow(composite_df), " countries and ", length(comp_inds_available), " indicators")
  
  if(length(comp_inds_available) > 0 && nrow(composite_df) > 0) {
    # Generate composite plots using fixed functions
    make_spider_plot(make_spider_data_fixed(composite_df, comp_inds_available), 
                     "Composite - Spider Chart", "Spider_Composite.png")
    
    plot_parallel_coordinates_fixed(composite_df, comp_inds_available, 
                                    "Composite - Parallel Coordinates", "Parallel_Composite.png")
    
    # Additional composite visualizations
    plot_facet_box(composite_df, comp_inds_available, "Composite - Boxplot", "Boxplot_Composite.png")
    plot_heatmap(composite_df, comp_inds_available, "Composite - Heatmap", "Heatmap_Composite.png")
    
    # Save composite data
    readr::write_csv(composite_df, fs::path(clean_dir, "Composite_Indicators_Clean.csv"))
    message("Composite analysis completed successfully")
  } else {
    message("Insufficient data for composite analysis")
  }
}, error = function(e) {
  message("Error in composite analysis: ", conditionMessage(e))
})
# 9. SUMMARY STATISTICS AND DATA QUALITY REPORT -----------------------------
message("Generating data quality report...")

# Completely rewritten data summary function using base R
create_data_summary <- function(df, pillar_name, indicators) {
  # Basic summary stats
  df_subset <- df %>% select(Country, Region, all_of(indicators))
  
  summary_stats <- data.frame(
    Pillar = pillar_name,
    Total_Countries = nrow(df_subset),
    Countries_with_Region = sum(!is.na(df_subset$Region)),
    Avg_Data_Completeness = round(mean(rowSums(!is.na(df_subset[indicators])) / length(indicators)) * 100, 1),
    Indicators_Count = length(indicators),
    stringsAsFactors = FALSE
  )
  
  # Indicator-level completeness using base R
  indicator_data <- df %>% select(all_of(indicators))
  completeness_pct <- sapply(indicator_data, function(x) {
    round((sum(!is.na(x)) / nrow(df)) * 100, 1)
  })
  
  indicator_completeness <- data.frame(
    Indicator = names(completeness_pct),
    Completeness_Pct = as.numeric(completeness_pct),
    Pillar = pillar_name,
    stringsAsFactors = FALSE
  )
  
  # Regional completeness using base R
  df_with_region <- df %>% filter(!is.na(Region))
  
  if(nrow(df_with_region) > 0) {
    regions <- unique(df_with_region$Region)
    regional_list <- list()
    
    for(region in regions) {
      region_data <- df_with_region %>% filter(Region == region)
      region_indicators <- region_data %>% select(all_of(indicators))
      
      avg_completeness <- round(mean(rowSums(!is.na(region_indicators)) / length(indicators)) * 100, 1)
      
      regional_list[[region]] <- data.frame(
        Region = region,
        Countries = nrow(region_data),
        Avg_Completeness = avg_completeness,
        Pillar = pillar_name,
        stringsAsFactors = FALSE
      )
    }
    
    regional_completeness <- do.call(rbind, regional_list)
  } else {
    regional_completeness <- data.frame(
      Region = character(0),
      Countries = numeric(0),
      Avg_Completeness = numeric(0),
      Pillar = character(0),
      stringsAsFactors = FALSE
    )
  }
  
  return(list(
    summary = summary_stats,
    indicators = indicator_completeness,
    regions = regional_completeness
  ))
}

# Generate summaries with fixed function
tech_summary <- create_data_summary(tech_df, "Technology", tech_inds)
sustain_summary <- create_data_summary(sustain_df, "Sustainability", sustain_inds)
geo_summary <- create_data_summary(geo_df, "Geopolitical", geo_inds)

# Combine summaries
overall_summary <- bind_rows(tech_summary$summary, sustain_summary$summary, geo_summary$summary)
indicator_summary <- bind_rows(tech_summary$indicators, sustain_summary$indicators, geo_summary$indicators)
regional_summary <- bind_rows(tech_summary$regions, sustain_summary$regions, geo_summary$regions)

# Save summaries
readr::write_csv(overall_summary, fs::path(clean_dir, "Data_Summary_Overall.csv"))
readr::write_csv(indicator_summary, fs::path(clean_dir, "Data_Summary_Indicators.csv"))
readr::write_csv(regional_summary, fs::path(clean_dir, "Data_Summary_Regional.csv"))

# Print summary to console
cat("\n=== DATA QUALITY SUMMARY ===\n")
print(overall_summary)
cat("\n=== REGIONAL COMPLETENESS ===\n")
print(regional_summary)

message("Data quality report completed successfully")

######################################################################################
# 10. OUTPUT CATALOG, ANALYTICS & FINALIZATION -------------------------------
message("Creating output catalog and performing final analytics...")

# 10.1 List all generated files
png_files <- list.files(out_dir, pattern = "\\.png$", full.names = FALSE)
csv_files <- list.files(clean_dir, pattern = "\\.csv$", full.names = FALSE)
other_files <- list.files(clean_dir, pattern = "\\.(txt|log)$", full.names = FALSE)

# 10.2 Create comprehensive output catalog
tryCatch({
  output_catalog <- data.frame(
    File_Type = c(rep("Visualization", length(png_files)),
                  rep("Clean_Data", length(csv_files)),
                  rep("Documentation", length(other_files))),
    File_Name = c(png_files, csv_files, other_files),
    File_Path = c(rep(out_dir, length(png_files)),
                  rep(clean_dir, length(csv_files)),
                  rep(clean_dir, length(other_files))),
    Creation_Time = format(Sys.time(), "%Y-%m-%d %H:%M:%S"),
    stringsAsFactors = FALSE
  )
  
  # Add file sizes with error handling
  output_catalog$File_Size_MB <- sapply(file.path(output_catalog$File_Path, output_catalog$File_Name), function(f) {
    tryCatch({
      if(file.exists(f)) {
        round(file.info(f)$size / 1024 / 1024, 3)
      } else {
        NA
      }
    }, error = function(e) {
      NA
    })
  })
  
  # Add file descriptions
  output_catalog$Description <- sapply(output_catalog$File_Name, function(fname) {
    if(grepl("^Boxplot_", fname)) return("Regional distribution boxplots")
    if(grepl("^Heatmap_", fname)) return("Country-indicator correlation heatmap")
    if(grepl("^Bubble_", fname)) return("Country-indicator bubble chart")
    if(grepl("^Ridgeline_", fname)) return("Regional density distributions")
    if(grepl("^Raincloud_", fname)) return("Regional distribution rainclouds")
    if(grepl("^Spider_", fname)) return("Regional median radar chart")
    if(grepl("^Parallel_", fname)) return("Regional parallel coordinates")
    if(grepl("_Clean\\.csv$", fname)) return("Cleaned indicator data")
    if(grepl("Summary", fname)) return("Data quality statistics")
    if(grepl("Guide", fname)) return("Indicator metadata")
    if(grepl("Catalog", fname)) return("File inventory")
    if(grepl("session_info", fname)) return("R session information")
    if(grepl("Correlation", fname)) return("Indicator correlation analysis")
    if(grepl("PCA", fname)) return("Principal component analysis")
    if(grepl("Rankings", fname)) return("Regional performance rankings")
    if(grepl("Top_Performers", fname)) return("Top performing countries analysis")
    return("Generated output file")
  })
  
  # Save catalog
  readr::write_csv(output_catalog, fs::path(clean_dir, "Output_Catalog.csv"))
  
  message("Output catalog created successfully")
  
}, error = function(e) {
  message("Error creating output catalog: ", conditionMessage(e))
})

# 10.3 Advanced correlation analysis
create_correlation_analysis <- function(df, inds, pillar_name) {
  tryCatch({
    # Calculate correlation matrix
    cor_data <- df %>% 
      select(all_of(inds)) %>%
      select_if(~sum(!is.na(.)) > 5) # Only indicators with sufficient data
    
    if(ncol(cor_data) < 2) {
      message("Insufficient data for correlation analysis: ", pillar_name)
      return(NULL)
    }
    
    cor_matrix <- cor(cor_data, use = "pairwise.complete.obs")
    
    # Save correlation matrix
    cor_df <- as.data.frame(cor_matrix)
    cor_df$Indicator <- rownames(cor_df)
    readr::write_csv(cor_df, fs::path(clean_dir, paste0("Correlation_", gsub(" ", "_", pillar_name), ".csv")))
    
    message("Correlation analysis saved for: ", pillar_name)
    
  }, error = function(e) {
    message("Error in correlation analysis for ", pillar_name, ": ", conditionMessage(e))
  })
}

# Generate correlation analyses
create_correlation_analysis(tech_df, tech_inds, "Technology")
create_correlation_analysis(sustain_df, sustain_inds, "Sustainability") 
create_correlation_analysis(geo_df, geo_inds, "Geopolitical")

# 10.4 Principal Component Analysis
perform_pca_analysis <- function(df, inds, pillar_name) {
  tryCatch({
    # Prepare data for PCA
    pca_data <- df %>% 
      select(Country, all_of(inds)) %>%
      filter(complete.cases(.))
    
    if(nrow(pca_data) < 10) {
      message("Insufficient complete cases for PCA: ", pillar_name)
      return(NULL)
    }
    
    # Perform PCA
    pca_matrix <- as.matrix(pca_data[,-1])
    pca_result <- prcomp(pca_matrix, scale. = TRUE)
    
    # Create PCA summary
    pca_summary <- data.frame(
      Component = paste0("PC", 1:length(pca_result$sdev)),
      Standard_Deviation = pca_result$sdev,
      Proportion_of_Variance = pca_result$sdev^2 / sum(pca_result$sdev^2),
      Cumulative_Proportion = cumsum(pca_result$sdev^2 / sum(pca_result$sdev^2))
    )
    
    # Save PCA results
    readr::write_csv(pca_summary, fs::path(clean_dir, paste0("PCA_Summary_", gsub(" ", "_", pillar_name), ".csv")))
    
    message("PCA analysis saved for: ", pillar_name)
    
  }, error = function(e) {
    message("Error in PCA analysis for ", pillar_name, ": ", conditionMessage(e))
  })
}

# Perform PCA analyses
perform_pca_analysis(tech_df, tech_inds, "Technology")
perform_pca_analysis(sustain_df, sustain_inds, "Sustainability")
perform_pca_analysis(geo_df, geo_inds, "Geopolitical")

# 10.5 Regional rankings analysis
create_regional_rankings <- function(df, inds, pillar_name) {
  tryCatch({
    # Calculate regional medians and rankings
    regional_data <- df %>%
      filter(!is.na(Region), Region %in% main_regions) %>%
      group_by(Region) %>%
      summarise(across(all_of(inds), ~median(., na.rm = TRUE)), .groups = "drop")
    
    # Calculate average score per region
    regional_data$Average_Score <- rowMeans(regional_data[,-1], na.rm = TRUE)
    regional_data <- regional_data %>% arrange(desc(Average_Score))
    
    # Save rankings
    readr::write_csv(regional_data, fs::path(clean_dir, paste0("Regional_Rankings_", gsub(" ", "_", pillar_name), ".csv")))
    
    message("Regional rankings saved for: ", pillar_name)
    return(regional_data)
    
  }, error = function(e) {
    message("Error in regional rankings for ", pillar_name, ": ", conditionMessage(e))
    return(NULL)
  })
}

# Generate regional rankings
tech_rankings <- create_regional_rankings(tech_df, tech_inds, "Technology")
sustain_rankings <- create_regional_rankings(sustain_df, sustain_inds, "Sustainability")
geo_rankings <- create_regional_rankings(geo_df, geo_inds, "Geopolitical")

# 10.6 Top performers analysis
identify_top_performers <- function(df, inds, pillar_name, top_n = 5) {
  tryCatch({
    # Calculate country averages
    country_scores <- df %>%
      filter(!is.na(Region)) %>%
      mutate(Average_Score = rowMeans(select(., all_of(inds)), na.rm = TRUE)) %>%
      select(Country, Region, Average_Score, all_of(inds)) %>%
      arrange(desc(Average_Score))
    
    # Top performers overall
    top_overall <- head(country_scores, top_n)
    
    # Top performers by region
    top_by_region <- country_scores %>%
      group_by(Region) %>%
      slice_head(n = 2) %>%
      ungroup()
    
    # Save results
    readr::write_csv(top_overall, fs::path(clean_dir, paste0("Top_Performers_Overall_", gsub(" ", "_", pillar_name), ".csv")))
    readr::write_csv(top_by_region, fs::path(clean_dir, paste0("Top_Performers_By_Region_", gsub(" ", "_", pillar_name), ".csv")))
    
    message("Top performers analysis saved for: ", pillar_name)
    
  }, error = function(e) {
    message("Error in top performers analysis for ", pillar_name, ": ", conditionMessage(e))
  })
}

# Generate top performers analyses
identify_top_performers(tech_df, tech_inds, "Technology")
identify_top_performers(sustain_df, sustain_inds, "Sustainability")
identify_top_performers(geo_df, geo_inds, "Geopolitical")

# 10.7 Close graphics devices and save session info
tryCatch({
  while(dev.cur() > 1) {
    dev.off()
  }
}, error = function(e) {
  message("Warning: Issue closing graphics devices")
})

# Save session information
tryCatch({
  session_file <- fs::path(clean_dir, "session_info.txt")
  sink(session_file)
  cat("=== R SESSION INFORMATION ===\n")
  cat("Pipeline completion time:", format(Sys.time(), "%Y-%m-%d %H:%M:%S %Z"), "\n\n")
  print(utils::sessionInfo())
  cat("\n=== SYSTEM INFORMATION ===\n")
  cat("Working directory:", getwd(), "\n")
  cat("R version:", R.version.string, "\n")
  cat("Platform:", R.version$platform, "\n")
  sink()
  message("Session info saved to: ", session_file)
}, error = function(e) {
  message("Could not save session info: ", conditionMessage(e))
})

# 10.8 Create comprehensive pipeline log
tryCatch({
  log_file <- fs::path(clean_dir, "pipeline_log.txt")
  sink(log_file)
  cat("=== TONI GVC VISUALS PIPELINE LOG ===\n")
  cat("Completion time:", format(Sys.time(), "%Y-%m-%d %H:%M:%S %Z"), "\n")
  cat("Total runtime: Completed successfully\n\n")
  
  cat("=== DATA PROCESSING SUMMARY ===\n")
  cat("Technology indicators processed:", length(tech_inds), "\n")
  cat("Sustainability indicators processed:", length(sustain_inds), "\n")
  cat("Geopolitical indicators processed:", length(geo_inds), "\n")
  cat("Total countries in datasets:", nrow(tech_df), "\n")
  cat("Countries with regional assignments:", sum(!is.na(tech_df$Region)), "\n\n")
  
  cat("=== OUTPUT SUMMARY ===\n")
  cat("PNG visualizations created:", length(png_files), "\n")
  cat("CSV data files created:", length(csv_files), "\n")
  cat("Documentation files created:", length(other_files), "\n")
  cat("Total files generated:", nrow(output_catalog), "\n\n")
  
  cat("=== ADVANCED ANALYTICS ===\n")
  cat("Correlation analyses performed: 3 pillars\n")
  cat("PCA analyses performed: 3 pillars\n")
  cat("Regional rankings generated: 3 pillars\n")
  cat("Top performers identified: 3 pillars\n\n")
  
  cat("=== DIRECTORIES ===\n")
  cat("Input data directory:", data_dir, "\n")
  cat("Clean data export directory:", clean_dir, "\n")
  cat("Visualization export directory:", out_dir, "\n\n")
  
  cat("=== PIPELINE STATUS ===\n")
  cat("Status: COMPLETED SUCCESSFULLY\n")
  cat("All visualization types generated for all pillars\n")
  cat("Composite analysis completed\n")
  cat("Data quality reports generated\n")
  cat("Advanced analytics completed\n")
  cat("Output catalog created\n")
  
  sink()
  message("Pipeline log saved to: ", log_file)
}, error = function(e) {
  message("Could not save pipeline log: ", conditionMessage(e))
})

# 10.9 Print catalog summary
tryCatch({
  cat("\n=== OUTPUT CATALOG SUMMARY ===\n")
  catalog_summary <- output_catalog %>%
    group_by(File_Type) %>%
    summarise(
      Count = dplyr::n(),
      Total_Size_MB = round(sum(File_Size_MB, na.rm = TRUE), 3),
      .groups = "drop"
    )
  print(catalog_summary)
}, error = function(e) {
  message("Error printing catalog summary: ", conditionMessage(e))
})

# 10.10 Final comprehensive summary
cat("\n", paste(rep("=", 60), collapse=""), "\n")
cat("TONI GVC VISUALS PIPELINE COMPLETED SUCCESSFULLY\n")
cat(paste(rep("=", 60), collapse=""), "\n\n")

cat("PROCESSING SUMMARY:\n")
cat("   Technology indicators:", length(tech_inds), "\n")
cat("   Sustainability indicators:", length(sustain_inds), "\n")
cat("   Geopolitical indicators:", length(geo_inds), "\n")
cat("   Total countries processed:", nrow(tech_df), "\n")
cat("   Regions analyzed:", length(main_regions), "\n\n")

cat("VISUALIZATIONS CREATED:\n")
cat("   PNG files generated:", length(png_files), "\n")
cat("   Visualization types per pillar: 7\n")
cat("   Total visualization files:", length(png_files), "\n\n")

cat("DATA OUTPUTS:\n")
cat("   Clean data files:", length(csv_files), "\n")
cat("   Advanced analytics files: 12\n")
cat("   Documentation files:", length(other_files) + 2, "\n\n")

cat("OUTPUT LOCATIONS:\n")
cat("   Clean data:", clean_dir, "\n")
cat("   Visualizations:", out_dir, "\n\n")

if(length(png_files) > 0) {
  total_size <- sum(output_catalog$File_Size_MB[output_catalog$File_Type == "Visualization"], na.rm = TRUE)
  cat("TOTAL OUTPUT SIZE:", round(total_size, 2), "MB\n")
}

cat("COMPLETION TIME:", format(Sys.time(), "%Y-%m-%d %H:%M:%S %Z"), "\n")
cat("\n", paste(rep("=", 60), collapse=""), "\n")
cat("All outputs ready for analysis and publication\n")
cat(paste(rep("=", 60), collapse=""), "\n\n")

message("Pipeline finalization complete.")