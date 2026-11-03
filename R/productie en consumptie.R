library(tidyverse)
library(readxl)

datadir <- "C:/Users/MartinPastoors/Martin Pastoors/WKC - Voedseluitzee/Data/Nutrienten_uit_zee"

# Read the sheet and tidy the names
df <- read_excel(file.path(datadir, "produced and consumed in NL.xlsx")) |>
  rename(type = Type, soort = Soort, g = `Availability per capita (g/week)`) |>
  drop_na(g) |>
  filter(soort != "Overig") |>
  mutate(type = recode(type,
                       "Produced in NL diet" = "Geproduceerd in Nederland",
                       "Consumed in NL diet" = "Geconsumeerd in Nederland (2023)"),
         type = factor(type, levels = c("Geproduceerd in Nederland",
                                        "Geconsumeerd in Nederland (2023)")))

# Soortgroepen: soorten + kleurverloop (donker -> licht)
groepen <- list(
  "Pelagische vis"                = list(c("Haring", "Blauwe wijting", "Makreel", "Horsmakreel",
                                           "Sardien", "Tonijn"),             c("#1f5f8b", "#9ecae1")),
  "Demersale vis"  = list(c("Alaska koolvis", "Kabeljauw", "Zwarte koolvis",
                                           "Schol", "Tong"),                 c("#2d7d46", "#a8d5a2")),
  "Kweekvis"                      = list(c("Zalm", "Pangasius", "Tilapia"),  c("#6a4c93", "#c9b6e4")),
  "Schaaldieren, inktvis" = list(c("Mosselen", "Schelpen", "Garnaal",
                                              "Tropische garnaal", "Inktvis"), c("#c4661f", "#f6c48f"))
)

# Kleur per soort
pal_soort <- groepen |>
  map(\(x) setNames(colorRampPalette(x[[2]])(length(x[[1]])), x[[1]])) |>
  unname() |> unlist() 
  # c(Overig = "grey75")

# Kleur per groep (middentint van elke familie) voor de legenda
pal_groep <- groepen |>
  map_chr(\(x) colorRampPalette(x[[2]])(3)[2]) 
  # c(Overig = "grey75")

groep_tbl <- imap_dfr(groepen, \(x, nm) tibble(soort = x[[1]], groep = nm))

# Volgorde per paneel, groep koppelen, labels
df_plot <- df |>
  left_join(groep_tbl, by = "soort") |>
  mutate(groep = factor(replace_na(groep, "Overig"), levels = names(pal_groep))) |>
  group_by(type) |>
  arrange(soort == "Overig", desc(g), .by_group = TRUE) |>
  mutate(volgorde = factor(row_number())) |>
  ungroup() |>
  mutate(label_groot = if_else(g >= 10, soort, ""),
         label_klein = if_else(g >= 1 & g < 10, soort, ""))

ggplot(df_plot, aes(x = g, y = "", fill = soort, group = volgorde)) +
  geom_col(position = position_stack(reverse = TRUE), width = 0.8,
           colour = "white", linewidth = 0.3) +
  geom_text(aes(label = label_groot),
            position = position_stack(reverse = TRUE, vjust = 0.5),
            angle = 90, size = 3.2, fontface = "bold", colour = "grey10") +
  geom_text(aes(label = label_klein),
            position = position_stack(reverse = TRUE, vjust = 0.5),
            angle = 90, size = 2.2, colour = "grey10") +
  # onzichtbare laag, alleen voor de groepslegenda
  # geom_point(aes(x = 0, colour = groep), size = 0, alpha = 0, inherit.aes = FALSE,
  #            data = df_plot |> mutate(y = "")) +
  geom_point(aes(x = 0, y = "", colour = groep), size = 0, alpha = 0, inherit.aes = FALSE) +
  geom_vline(xintercept = 100, linetype = "dashed", linewidth = 0.8) +
  annotate("text", x = 101, y = 1.48, hjust = 0, size = 3,
           label = "Richtlijn Voedingscentrum (2023): 100 g per week") +
  facet_wrap(~ type, ncol = 1) +
  scale_fill_manual(values = pal_soort, guide = "none") +
  scale_colour_manual(values = pal_groep, name = NULL, drop = FALSE) +
  guides(colour = guide_legend(nrow = 1,
                               override.aes = list(shape = 15, size = 6, alpha = 1))) +
  scale_x_continuous(limits = c(0, 170), breaks = seq(0, 160, 20), expand = c(0, 0)) +
  labs(x = "Vis en schaal- en schelpdieren (gram per persoon per week)", y = NULL) +
  theme_minimal(base_size = 12) +
  theme(panel.grid.major.y = element_blank(),
        panel.grid.minor   = element_blank(),
        strip.text         = element_text(face = "bold", hjust = 0),
        legend.position    = "bottom",
        legend.text        = element_text(size = 10))

ggsave("vis_NL_productie_vs_consumptie.png", width = 10, height = 6.5, dpi = 300)
