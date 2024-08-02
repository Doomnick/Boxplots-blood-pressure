library(readxl)
library(tidyr)
library(dplyr)
library(readxl)
library(dplyr)
library(afex) 
library(emmeans)
library(writexl)
library(purrr)

# Load the Excel sheet into R

file_path <- "C:/Users/***/Desktop/***/15.7.2024/Formát pro série.xlsx"
sheet_names <- excel_sheets(file_path)

i<-1
for (i in 1:length(sheet_names)) {
sheet_name <- sheet_names[i]

# Read the data from the specified sheet
data <- read_excel(file_path, sheet = sheet_name)


data <- data %>%
  mutate(
    Subject = factor(Subject),
    Group = factor(Group),
    Sex = factor(Sex),
    Intervention = factor(Intervention),
    Exercise = factor(Exercise),
    Set = factor(Set)
  )


anova_sys <- aov_ez(
  id = "Subject",               
  dv = "Sys",                   
  data = data,                
  within = c("Set"),            
  between = c("Group"),       
  anova_table = list(correction = "GG") 
)



anova_dia <- aov_ez(
  id = "Subject",                
  dv = "Dia",                   
  data = data,                  
  within = c("Set"),            
  between = c("Group"),       
  anova_table = list(correction = "GG") 
)

anova_pwvao <- aov_ez(
  id = "Subject",                
  dv = "PWVao",                  
  data = data,                   
  within = c("Set"),             
  between = c("Group"),      
  anova_table = list(correction = "GG") 
)


posthoc_sys <- emmeans(anova_sys, pairwise ~ Set | Group, adjust = "bonferroni")
summary(posthoc_sys)


posthoc_dia <- emmeans(anova_dia, pairwise ~ Set | Group, adjust = "bonferroni")
summary(posthoc_dia)


posthoc_pwvao <- emmeans(anova_pwvao, pairwise ~ Set | Group, adjust = "bonferroni")
summary(posthoc_pwvao)


anova_sys_summary <- as.data.frame(anova_sys$anova_table)
anova_sys_summary$Term <- rownames(anova_sys_summary)  

anova_dia_summary <- as.data.frame(anova_dia$anova_table)
anova_dia_summary$Term <- rownames(anova_dia_summary)  

anova_pwvao_summary <- as.data.frame(anova_pwvao$anova_table)
anova_pwvao_summary$Term <- rownames(anova_pwvao_summary)

anova_sys_summary <- anova_sys_summary %>%
  relocate(Term, .before = everything())

anova_dia_summary <- anova_dia_summary %>%
  relocate(Term, .before = everything())

anova_pwvao_summary <- anova_pwvao_summary %>%
  relocate(Term, .before = everything())



posthoc_sys_summary <- as.data.frame(summary(posthoc_sys)$contrasts)
posthoc_dia_summary <- as.data.frame(summary(posthoc_dia)$contrasts)
posthoc_pwvao_summary <- as.data.frame(summary(posthoc_pwvao)$contrasts)

colnames(posthoc_sys_summary) <- c("Comparison", "Group", "Estimate", "SE", "df", "t-value", "p-value")
colnames(posthoc_dia_summary) <- c("Comparison", "Group", "Estimate", "SE", "df", "t-value", "p-value")
colnames(posthoc_pwvao_summary) <- c("Comparison", "Group", "Estimate", "SE", "df", "t-value", "p-value")


list_of_results <- list(
  "ANOVA Sys" = anova_sys_summary,
  "ANOVA Dia" = anova_dia_summary,
  "ANOVA PWVao" = anova_pwvao_summary,
  "Posthoc Sys" = posthoc_sys_summary,
  "Posthoc Dia" = posthoc_dia_summary,
  "Posthoc PWVao" = posthoc_pwvao_summary
)




write_xlsx(
  list(
    "ANOVA Sys" = anova_sys_summary,
    "ANOVA Dia" = anova_dia_summary,
    "ANOVA PWVao" = anova_pwvao_summary,
    "Posthoc Sys" = posthoc_sys_summary,
    "Posthoc Dia" = posthoc_dia_summary,
    "Posthoc PWVao" = posthoc_pwvao_summary
  ),
  path = paste("C:/Users/***/Desktop/***/15.7.2024/", sheet_name, ".xlsx", sep=""),
) 

}

