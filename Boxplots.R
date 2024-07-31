library(ggplot2)
library(dplyr)
library(readxl)
library(gridExtra)
library(grid)

setwd("C:/Users/***/Desktop/***/15.7.2024")
sheets <- excel_sheets("C:/Users/***/Desktop/***/15.7.2024/Formát pro série.xlsx")

file_path <- "C:/Users/***r/Desktop/***/15.7.2024/Formát pro série.xlsx"
for (i in 1:length(sheets)) {
  data <- read_excel(file_path, sheet = sheets[i])
  
  data <- data %>%
    mutate(Exercise_Set_Group = ifelse(Set == "BV", Exercise, paste(Exercise, "Set", Set, Group)),
           Exercise = factor(Exercise, levels = unique(Exercise)))
  

  data <- data %>%
    arrange(Exercise, Group, Set)
  
  remove_outliers <- function(df) {
    Q1 <- quantile(df$PWVao, 0.25, na.rm = TRUE)
    Q3 <- quantile(df$PWVao, 0.75, na.rm = TRUE)
    IQR <- Q3 - Q1
    df %>%
      filter(PWVao >= (Q1 - 3 * IQR) & PWVao <= (Q3 + 3 * IQR))
  }
  
  data <- data %>%
    group_by(Exercise, Group, Set) %>%
    do(remove_outliers(.)) %>%
    ungroup()
  

  exercises <- unique(data$Exercise)
  
  y_limits <- range(data$PWVao, na.rm = TRUE)
  

  create_plot <- function(exercise_data, exercise_name, plot_number, total_plots) {
    plot <- ggplot(exercise_data, aes(x = Set, y = PWVao, fill = Group)) +
      geom_boxplot(position = position_dodge(width = 0.8), width = 0.5) +
      labs(title = exercise_name,
           x = if (plot_number == 3) "Set" else "",  # Conditionally set x-axis label
           y = "Aortic Pulse Wave (m/s)") +  
      theme(axis.text.x = element_text(angle = 45, hjust = 1, size = 8),
            strip.text = element_text(size = 12),
            plot.title = element_text(hjust = 0.5, size = 14)) +  
      scale_fill_manual(values = c("stage I hypertension" = "red", "normotension" = "blue")) +
      theme_classic() +
      scale_y_continuous(limits = y_limits, breaks = seq(floor(y_limits[1] / 10) * 10, ceiling(y_limits[2] / 10) * 10, by = 10)) +
      theme(
        panel.grid.major = element_line(color = "grey90", size = 0.3, linetype = "solid"),
        panel.grid.minor = element_line(color = "grey90", size = 0.3, linetype = "solid"),
        axis.title.y = element_blank(),
        legend.position = "none"
      )
    
    if (plot_number != 1) {
      plot <- plot +
        theme_minimal() + 
        theme(
          axis.text.y = element_blank(),
          axis.ticks.y = element_blank(),
          axis.line.x = element_line(),
          axis.title.y = element_blank(),
          legend.position = "none"
        ) 
      if (plot_number == 3) {
        plot <- plot + labs(x = "Set")
      }
    }
    
    return(plot)
  }
  

  get_legend <- function(my_plot) {
    tmp <- ggplot_gtable(ggplot_build(my_plot))
    leg <- which(sapply(tmp$grobs, function(x) x$name) == "guide-box")
    legend <- tmp$grobs[[leg]]
    return(legend)
  }
  

  plots <- list()
  

  for (j in 1:length(exercises)) {
    exercise_data <- data %>% filter(Exercise == exercises[j])
    plot <- create_plot(exercise_data, exercises[j], plot_number = j, total_plots = length(exercises))
    plots[[j]] <- plot
  }
  

  legend_plot <- ggplot(data %>% filter(Exercise == exercises[1]), aes(x = Set, y = PWVao, fill = Group)) +
    geom_boxplot(position = position_dodge(width = 0.8), width = 0.5) +
    scale_fill_manual(values = c("stage I hypertension" = "red", "normotension" = "blue")) +
    theme_classic() +
    theme(legend.position = "bottom")
  

  legend <- get_legend(legend_plot)
  

  combined_plot <- grid.arrange(
    grobs = plots, ncol = length(plots),
    top = paste(sheets[i]," - Aortic Pulse Wave by Exercise and Set"),
    left = textGrob("Aortic Pulse Wave (m/s)", rot = 90, vjust = 1, gp = gpar(cex = 1.3)),
    bottom = legend
  )
  
  file_name <- paste0("PWVao_", sheets[i], ".png")
  ggsave(file_name, combined_plot, width = 1000/96, height = 539/96, dpi = 96, units = "in")
  
}
