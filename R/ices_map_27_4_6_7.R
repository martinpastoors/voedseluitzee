# =====================================================================
# Map of ICES subareas 27.4, 27.6 and 27.7 (Northeast Atlantic)
# with ICES divisions, EEZ boundaries and country outlines
# =====================================================================
# Data sources (downloaded automatically on first run, then cached):
#   - ICES statistical areas : ICES GIS (gis.ices.dk)
#   - EEZ boundaries         : Marine Regions / Flanders Marine Institute
#                              (via the mregions2 package)
#   - Countries              : Natural Earth (via rnaturalearth)
# =====================================================================

# ---- 0. Packages ----------------------------------------------------
pkgs <- c("sf", "dplyr", "ggplot2", "rnaturalearth", "mregions2")
to_install <- pkgs[!pkgs %in% rownames(installed.packages())]
if (length(to_install)) install.packages(to_install)

library(sf)
library(dplyr)
library(ggplot2)
library(rnaturalearth)
library(mregions2)

sf_use_s2(FALSE)                     # planar operations: fewer geometry errors
options(timeout = max(600, getOption("timeout")))

data_dir <- "map_data"
dir.create(data_dir, showWarnings = FALSE)

subareas_wanted <- c("4", "6", "7")
crs_map <- 3035                      # ETRS89 / LAEA Europe (equal-area)

# ---- 1. ICES statistical areas ---------------------------------------
ices_zip <- file.path(data_dir, "ICES_areas.zip")
ices_dir <- file.path(data_dir, "ICES_areas")

if (!file.exists(ices_zip)) {
  download.file("https://gis.ices.dk/shapefiles/ICES_areas.zip",
                ices_zip, mode = "wb")
}
if (!dir.exists(ices_dir)) unzip(ices_zip, exdir = ices_dir)

ices_shp <- list.files(ices_dir, pattern = "\\.shp$",
                       full.names = TRUE, recursive = TRUE)[1]

ices <- st_read(ices_shp, quiet = TRUE) |>
  st_make_valid() |>
  st_transform(4326)

# 'Area_27' holds codes such as "4.a", "6.b.2", "7.j.1".
# Derive subarea ("7") and division ("7.j") from it.
ices_sel <- ices |>
  mutate(code     = as.character(Area_27),
         subarea  = sub("^(\\d+).*", "\\1", code),
         division = sub("^(\\d+\\.[a-z]+).*", "\\1", code)) |>
  filter(subarea %in% subareas_wanted)

# Dissolve sub-divisions into divisions, and divisions into subareas
divisions <- ices_sel |>
  group_by(subarea, division) |>
  summarise(.groups = "drop") |>
  st_make_valid()

subareas <- ices_sel |>
  group_by(subarea) |>
  summarise(.groups = "drop") |>
  st_make_valid()

# Study area: subarea 27.4 (North Sea) + division 27.7.d (Eastern Channel)
study_area <- divisions |>
  filter(subarea == "4" | division == "7.d") |>
  summarise() |>
  st_make_valid()

# Label points (computed in the map projection so they sit inside polygons)
div_labels <- divisions |>
  st_transform(crs_map) |>
  st_point_on_surface() |>
  mutate(label = paste0("27.", division))

# ---- 2. Map extent ---------------------------------------------------
bb_map <- st_bbox(st_transform(subareas, crs_map))
pad    <- 100000                     # 100 km margin
xlim   <- c(bb_map["xmin"] - pad, bb_map["xmax"] + pad)
ylim   <- c(bb_map["ymin"] - pad, bb_map["ymax"] + pad)

# Generous lon/lat box for cropping the background layers
bb_ll  <- st_bbox(subareas)
crop_box <- st_bbox(c(xmin = bb_ll[["xmin"]] - 8, xmax = bb_ll[["xmax"]] + 8,
                      ymin = bb_ll[["ymin"]] - 5, ymax = bb_ll[["ymax"]] + 5),
                    crs = st_crs(4326))

# ---- 3. Countries ----------------------------------------------------
# scale = "large" gives finer coastlines but needs the rnaturalearthhires package
countries <- ne_countries(scale = "medium", returnclass = "sf") |>
  st_make_valid() |>
  st_crop(crop_box)

# ---- 4. EEZ boundaries (Marine Regions) ------------------------------
eez_file <- file.path(data_dir, "eez_boundaries.rds")
if (!file.exists(eez_file)) {
  saveRDS(mrp_get("eez_boundaries"), eez_file)   # global line layer, cached
}
eez_lines <- readRDS(eez_file) |>
  st_make_valid() |>
  st_crop(crop_box)

# Dutch EEZ as a polygon, to highlight it (handy for a report on NL fisheries)
nl_eez_file <- file.path(data_dir, "eez_nld.rds")
if (!file.exists(nl_eez_file)) {
  saveRDS(mrp_get("eez", cql_filter = "iso_ter1 = 'NLD'"), nl_eez_file)
}
nl_eez <- readRDS(nl_eez_file) |> st_make_valid()

# ---- 5. Plot ---------------------------------------------------------
p <- ggplot() +
  # Dutch EEZ (fill only)
  geom_sf(data = nl_eez, aes(fill = "Dutch EEZ"), colour = NA) +
  # ICES divisions and the study area outline
  geom_sf(data = divisions, aes(colour = "ICES division"),
          fill = NA, linewidth = 0.35) +
  geom_sf(data = study_area, aes(colour = "Study area (27.4 + 27.7.d)"),
          fill = NA, linewidth = 1.1) +
  # EEZ boundaries
  geom_sf(data = eez_lines, aes(colour = "EEZ boundary"),
          linetype = "dashed", linewidth = 0.4) +
  # Land on top so area/EEZ lines don't run over the coast
  geom_sf(data = countries, fill = "grey88", colour = "grey45",
          linewidth = 0.2) +
  # Division labels
  geom_sf_text(data = div_labels, aes(label = label),
               size = 2.8, colour = "steelblue4", fontface = "bold") +
  scale_fill_manual(values = c("Dutch EEZ" = "#d6e8f5"), name = NULL) +
  scale_colour_manual(
    values = c("EEZ boundary"               = "firebrick",
               "ICES division"              = "steelblue3",
               "Study area (27.4 + 27.7.d)" = "navy"),
    name = NULL) +
  guides(colour = guide_legend(override.aes = list(
    linetype  = c("dashed", "solid", "solid"),
    linewidth = c(0.5, 0.4, 1.1)))) +
  coord_sf(crs = crs_map, xlim = xlim, ylim = ylim, expand = FALSE,
           default_crs = NULL) +
  labs(
    title    = "ICES subareas 27.4, 27.6 and 27.7",
    subtitle = "North Sea, West of Scotland and Celtic Seas / Western Channel",
    caption  = paste("Sources: ICES statistical areas (gis.ices.dk);",
                     "EEZ boundaries: Flanders Marine Institute, Marine Regions;",
                     "countries: Natural Earth."),
    x = NULL, y = NULL) +
  theme_minimal(base_size = 10) +
  theme(panel.background = element_rect(fill = "#f4f9fd", colour = NA),
        panel.grid       = element_line(colour = "white", linewidth = 0.3),
        legend.position  = "bottom",
        plot.caption     = element_text(size = 7, colour = "grey40"))

print(p)


ggsave("ices_27_4_6_7_map.png", p, width = 18, height = 20, units = "cm",
       dpi = 300, bg = "white")
