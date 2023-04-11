# ACTL4001 Assignment

data_dir <- "C:/Study/ACTL4001/Assignment/Data"

#### Setup ####
# Packages
packages <- c("openxlsx", "data.table", "ggplot2", "zoo", "corrplot",
              "gridExtra", "fitdistrplus", "actuar", "gamlss")
install.packages(packages)

library("openxlsx")
library("data.table")
library("ggplot2")
library("zoo")
library("corrplot")
library("gridExtra")
library("fitdistrplus")
library("actuar")
library("gamlss")

# Parameters defined below
long_term_inflation <- 0.020
temp_housing <- c(1920, 1829, 1925, 1639, 1588, 1653)
demand_push_ceiling <- 0.5
home_contents_min <- 0.4
home_contents_max <- 0.7
splits <- c("minor", "medium", "major")
# home_contents <- seq(0.4, 0.7, 0.05)
exc_rate <- 1.321

#### Data read ####

hazard <- read.xlsx(file.path(data_dir, "2023-student-research-hazard-event-data.xlsx"),
                    sheet = "Hazard Data",
                    rows = 13:3379, cols = 2:9)

dem <- read.xlsx(file.path(data_dir, "2023-student-research-eco-dem-data.xlsx"),
                 sheet = "Demographic-Economic",
                 rows = 8:51, cols = 2:8)

eco <- read.xlsx(file.path(data_dir, "2023-student-research-eco-dem-data.xlsx"),
                 sheet = "Inflation-Interest",
                 rows = 8:68, cols = 2:6)

raf <- read.xlsx(file.path(data_dir, "2023-student-research-emissions.xlsx"),
                 sheet = "Model",
                 rows = 20:34, cols = 20:23)

mmm <- read.xlsx(file.path(data_dir, "Minor, Medium, Major Risk.xlsx"),
                 sheet = "Summary Sheet w ranks",
                 rows = 4:55, cols = 16:19)

hazard <- as.data.table(hazard)
dem <- as.data.table(dem)
eco <- as.data.table(eco)
raf <- as.data.table(raf)
mmm <- as.data.table(mmm)

#### Data clean ####
hazard_edit <- copy(hazard)
hazard_edit[,Region := as.factor(Region)]

table(hazard_edit$Hazard.Event)
hazard_edit[Hazard.Event == "Severe Storm/Thunder Storm - Wind", Hazard.Event := "Severe Storm/Thunder Storm/ Wind"]


eco_edit <- copy(eco) #should use "copy" when duplicating/making a copy of data.tables
#Simple average interpolation for missing/erroneous rates *warning: ugly
eco_edit[Year == 2018,]$`1-yr.risk.free.rate` <- eco_edit[Year == 2018,]$Government.Overnight.Bank.Lending.Rate
eco_edit[Year == 2003,]$Inflation <- mean(c(eco_edit[Year == 2002,]$Inflation, eco_edit[Year == 2004,]$Inflation ))
eco_edit[Year == 1987,]$Government.Overnight.Bank.Lending.Rate <- eco_edit[Year == 2018,]$`1-yr.risk.free.rate`

#Extrapolate for 1960, 1961
eco_edit <- rbind(c("Year" = 1960, "Inflation" = long_term_inflation, eco_edit[Year == 1962][,3:5]),
                  c("Year" = 1961, "Inflation" = long_term_inflation, eco_edit[Year == 1962][,3:5]),
                  eco_edit)

#### Data prep ####
# Inflate payments
eco_edit <- eco_edit[order(Year, decreasing = TRUE),]
eco_edit[, inf_edit := Inflation + 1]
eco_edit[,cumulative_inf := cumprod(inf_edit)]

hazard_edit <- merge(hazard_edit, eco_edit[,.(Year, cumulative_inf)], by = "Year", all.x = TRUE)
hazard_edit[,prop_dam_inf := Property.Damage*cumulative_inf]

# Year and quarter
hazard_edit[, yq := as.yearqtr(paste0(Year,"-",Quarter))]

# Region as factor
hazard_edit[, Region := as.factor(Region)]

# Hazard flags
event_names <- unique(trimws(unlist(strsplit(unique(hazard_edit$Hazard.Event), "/"))))
for (i in 1:length(event_names)) {
  hazard_edit[, paste0("event_",sub(" ", "_",event_names[i])) := grepl(event_names[i], Hazard.Event, ignore.case = TRUE)]
}

#### Basic summaries ####
# Prop_dam_inf by region, by yq
ggplot(hazard_edit, aes(x = yq, y = prop_dam_inf, col = Region)) +
  geom_line() +
  ylim(0,1000000)

ggplot(hazard_edit, aes(x = Year, y = prop_dam_inf, col = Region)) +
  geom_line() +
  ylim(0,1000000)

ggplot(hazard_edit, aes(x = Year, col = Region)) +
  geom_line(stat = "count") +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1))

hazard_edit[,avg_dam_reg := mean(prop_dam_inf), by = c("Region", "yq")]

for (reg in c("1","2","3","4","5","6")) {
  dam_by_region_plot <- ggplot(hazard_edit[Region == reg,], aes(x = yq, y = prop_dam_inf)) +
    geom_line() +
    ylim(0,1000000)
  print(dam_by_region_plot)
}

for (reg in c("1","2","3","4","5","6")) {
  dam_by_region_plot <- ggplot(hazard_edit[Region == reg,], aes(x = yq, y = avg_dam_reg)) +
    geom_line() +
    ylim(0,1000000)
  print(dam_by_region_plot)
}

# Hazard by region, by yq
ggplot(hazard_edit, aes(x = Hazard.Event, col = Region)) +
  geom_bar() +
  theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1))

for (reg in c("1","2","3","4","5","6")) {
  hazard_by_region_plot <- ggplot(hazard_edit[Region == reg,], aes(x = Hazard.Event)) +
    geom_bar() +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1))
  print(hazard_by_region_plot)
}
# Region 3 has the most flooding, severe storm/thunderstorm/wind
# Region 2 has the most hail
# Region 1 has the most winter weather

# Correlation
correlation <- cor(hazard_edit[,.(Duration, Fatalities, Injuries, prop_dam_inf)])
corrplot(correlation)

event_index <- grepl("event", colnames(hazard_edit))
corr_hazards <- cor(hazard_edit[,..event_index])
corrplot(corr_hazards)

# Superimposed inflation (demand push/supply pull)
acs_by_freq <- hazard_edit[,.(.N, mean_acs = mean(prop_dam_inf)),by = yq]
corr_acs_by_freq <- cor(acs_by_freq[,.(N, mean_acs)])
corr_acs_by_freq #only 0.024

acs_by_freq_simple <- acs_by_freq[, .(mean_acs_again = mean(mean_acs)), by = N]
acs_by_freq_simple[order(N),]
plot(acs_by_freq_simple) #no clear relationship


# Prop_dam_inf by region, by hazard
plot_list <- list()
for (haz in event_names) {
  for (ind in c("1","2","3","4","5","6")) {
    reg <- ind
    plot_list[[ind]] <- ggplot(hazard_edit[Region == reg & get(paste0("event_", haz)) == TRUE ,], aes(x = prop_dam_inf)) +
      geom_histogram() +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1)) +
      ggtitle(paste0("Region ",reg," - ", haz))

  }
  grid.arrange(plot_list[[1]], plot_list[[2]], plot_list[[3]], plot_list[[4]], plot_list[[5]], plot_list[[6]], ncol = 3)
}

# Export for excel analysis
write.xlsx(hazard_edit[,.(Region, Hazard.Event, Quarter, Year, Duration, Fatalities, Injuries, prop_dam_inf)],
           file.path(data_dir, "inf_hazard_event_data.xlsx"),
           sheetName = "inf_data")


#### Minor/medium/major split ####
haz_mod_data <- copy(hazard_edit)

haz_mod_data <- merge(haz_mod_data, mmm[,.(Group, `Most.to.Least.Hazardous.(Based.on.weightings)`)], by.x = c("Hazard.Event"), by.y = c("Most.to.Least.Hazardous.(Based.on.weightings)"), all.x = TRUE)
haz_mod_data[is.na(Group)] #check - none!

haz_mod_data[, sum_hazard := prop_dam_inf]

haz_mod_data$post2005 <- 0
haz_mod_data[Year >= 2005, post2005 := 1]

#########################################
#### Predict number of events for each split, by region ####
# Check existing trends in notifications by split
## why are there ~ 0 claims ~ Year 2000?
## why is there a jump in claims in the recent 20 years?
## edit to definition of regions may explain reduction in region 1 notifications and sudden hike in other region notifications

# Frequency plots by year, by region and mmm group
plot_list <- list()
for (reg in c("1","2","3","4","5","6")) {
  plot_list[[1]] <- ggplot(haz_mod_data[Region == reg,], aes(x = Year)) +
    geom_bar(stat = "count") +
    theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1)) +
    ggtitle(paste0("Region ",reg), " count")

  for (ind in c(1,2,3)) {
    split <- splits[ind]
    plot_list[[ind + 1]] <- ggplot(haz_mod_data[Region == reg & Group == split,], aes(x = Year)) +
      geom_bar(stat = "count") +
      theme(axis.text.x = element_text(angle = 90, vjust = 0.5, hjust=1)) +
      ggtitle(paste0("Region ",reg," - ", split, " count"))
  }

  grid.arrange(plot_list[[1]], plot_list[[2]], plot_list[[3]], plot_list[[4]], ncol = 4)
}


# Confirm variation in experience increased considerably following 2005
sd(haz_mod_data[post2005 == FALSE, .N, yq]$N)
# [1] 10.39263
sd(haz_mod_data[post2005 == TRUE, .N, yq]$N) #lots more variation here
# [1] 24.60607
sd(haz_mod_data[, .N, yq]$N)
#[1] 15.89482

# Method 1 - Take average of last 10 years - 10 chosen arbitrarily
# freq_list <- list()
# for (reg in c("1","2","3","4","5","6")) {
#   reg_freq_list <- c()
#   for (group in splits) {
#     split_freq <- mean(haz_mod_data[Region == reg & Year >= 2010 & Group == group, .N, Year]$N)
#     reg_freq_list <- c(reg_freq_list, ifelse(is.na(split_freq),0, split_freq))
#   }
#   print(reg_freq_list)
#   freq_list[[reg]] <- reg_freq_list
# }
# write.xlsx(reg_freq_list,
#            file.path(data_dir, "events_group_freq.xlsx"),
#            sheetName = "freq")

# Method 2 - fit distribution
data_freq <- merge(haz_mod_data[sum_hazard > 0, .N, Year], data.table("Year" = seq(1960,2020,1)), by = c("Year"), all.y = TRUE)
data_freq[,N := ifelse(is.na(N), 0, N)]
data_freq.1 <- merge(haz_mod_data[post2005 == TRUE & sum_hazard > 0, .N, Year], data.table("Year" = seq(2006,2020,1)), by = c("Year"), all.y = TRUE)
data_freq.1[,N := ifelse(is.na(N), 0, N)]

freq_skeleton <- CJ(as.factor(1:6),splits, seq(1960,2020))
colnames(freq_skeleton) <- c("Region", "Group", "Year")
data_freq_groups <- merge(haz_mod_data[sum_hazard > 0, .N, c("Region", "Group", "Year")], freq_skeleton, by = c("Region", "Group", "Year"), all.y = TRUE)
data_freq_groups[,N := ifelse(is.na(N), 0, N)]
freq_skeleton.1 <- CJ(as.factor(1:6),splits, seq(2006,2020))
colnames(freq_skeleton.1) <- c("Region", "Group", "Year")
data_freq_groups.1 <- merge(haz_mod_data[post2005 == TRUE & sum_hazard > 0, .N, c("Region", "Group", "Year")], freq_skeleton.1, by = c("Region", "Group", "Year"), all.y = TRUE)
data_freq_groups.1[,N := ifelse(is.na(N), 0, N)]

# Distribution selected in ass_data_diagnostics.R
# Fit negative binomial distribution to notifications per year
## for each split and region
combined_data_freq <- rbind(data_freq_groups.1[Group == "major"],
                            data_freq_groups[Group == "medium"],
                            data_freq_groups[Group == "minor"])
freq_glm <-glm.nb(N ~ Region + Group, data = combined_data_freq)

summary(freq_glm)

freq_results <- predict(freq_glm,type = "response", se.fit = TRUE) #se.fit = TRUE is not supported for new data values at the moment
freq_results <- cbind(combined_data_freq, "fit_freq" = freq_results$fit, "fit_freq_se" = freq_results$se.fit)
freq_results_summ <- unique(freq_results[,.(Region, Group, fit_freq, fit_freq_se)])[order(Region, Group)]


# write.xlsx(freq_results_summ,
#            file.path(data_dir, "freq_sev_results.xlsx"),
#            sheetName = "tot_freq")

#########################################

#### Predict avg prop damage of events for each split, by region ####
# Not offset by populations of regions bc insufficient information provided
# We still assume a relationship between region and acs in projections.
event_names2 <- paste0("event_", sub(" ", "_", event_names))
colSums(haz_mod_data[,..event_names2])

# Only 1 event for fogs, landslides
# Severe storms and thunder storms exactly correlated.
# Hurricanes and tropical storms exactly correlated.
# Only hurricanes and tropical storms have significantly non-0 correlation with sum_hazard
event_names2_sh <- c(event_names2, "sum_hazard", "Quarter")
correlation <- cor(haz_mod_data[,..event_names2_sh])
corrplot(correlation)

# Distribution selected in ass_data_diagnostics.R
# Fit lognormal distribution to severities
# for each split and region
data_sev <- haz_mod_data[sum_hazard > 0,]
sev_glm <- gamlss(sum_hazard ~ Region + Group, family = LOGNO, link = identity, data = data_sev)
summary(sev_glm)

sev_results <- predict(sev_glm,type = "response", se.fit = TRUE) #se.fit = TRUE is not supported for new data values at the moment
sev_results <- cbind(data_sev, "fit_sev" = sev_results$fit, "fit_sev_se" = sev_results$se.fit)
# Cox method for confidence intervals https://jse.amstat.org/v13n1/olsson.html#:~:text=For%20a%2095%25%20confidence%20interval,(T2%3B0.975)%5D.
sev_results <- sev_results[data_sev[,.N, by = c("Region", "Group")], on = c("Region", "Group")]
sev_results[,fit_sev_sd := fit_sev_se*sqrt(N)]
sev_results[,`:=`(point_est = fit_sev + fit_sev_sd^2/2,
                  lower_bound = fit_sev + fit_sev_sd^2/2 - 2.02*sqrt(fit_sev_sd^2/N + fit_sev_sd^4/2/(N-1)),
                  upper_bound = fit_sev + fit_sev_sd^2/2 + 2.02*sqrt(fit_sev_sd^2/N + fit_sev_se^4/2/(N-1)))]
sev_results[,`:=`(trans_point_est = exp(point_est),
                  trans_lower_bound = exp(lower_bound),
                  trans_upper_bound = exp(upper_bound))]


sev_results_summ <- unique(sev_results[,.(Region, Group, trans_point_est, trans_lower_bound, trans_upper_bound)],
                           by = c("Region", "Group", "trans_point_est", "trans_lower_bound", "trans_upper_bound"))[order(Region, Group)]


# Export results
wb <- createWorkbook()
addWorksheet(wb, "tot_freq")
addWorksheet(wb, "nil_freq")
addWorksheet(wb, "pos_sev")

writeData(wb, "tot_freq", freq_results_summ, startRow = 1, startCol = 1)
writeData(wb, "nil_freq", nil_results_summ, startRow = 1, startCol = 1)
writeData(wb, "pos_sev", sev_results_summ, startRow = 1, startCol = 1)
saveWorkbook(wb, file = file.path(data_dir, "freq_sev_results.xlsx"), overwrite = TRUE)

#########################################

#########################################

#### Confidence analysis - for solvency ####
# Simulate variables - frequency
# Use sample mean of number of events across region and group for mu
library(dplyr)
data_freq_mu <- combined_data_freq %>%
  group_by(Region, Group) %>%
  summarise(mean_value = mean(N))

# Define hazard categories
regions <- unique(data_freq_mu$Region)
groups <- unique(data_freq_mu$Group)

# Create data frame to store results
mu_freq <- data.frame(region = rep(regions, each = length(groups)),
                      group = rep(groups, length(regions)),
                      mu = rep(NA, length(regions) * length(groups)))

# Input values for 'mu' and 'theta' into data frame
mu_freq$mu <- data_freq_mu$mean_value


# Set seed for reproducibility
set.seed(20)

# Simulate values
nsim <- 10000
i=1
freq_sim <- t(sapply(1:nrow(mu_freq), function(i)  rnbinom(nsim, mu = mu_freq$mu[i], size = freq_glm$theta)))
freq_sim <- data.frame(freq_sim)
colnames(freq_sim) <- paste0("sim_", 1:ncol(freq_sim))

# Combine simulated values with existing dataframe
freq_sim_combined <- cbind(mu_freq, freq_sim)
freq_sim_combined

# SEVERITY SIMULATION - Amanda
# Set up data frame
regions <- 1:6
groups <- c(unique(data_sev$Group))
set.seed(20) # for reproducibility

mu_int <- sev_glm$mu.coefficients[1]
mu_groups <- as.data.table(cbind("medium" = sev_glm$mu.coefficients[7], "minor" = sev_glm$mu.coefficients[8]))
# Create data frame
mu_coeff <- data.frame(region = rep(regions, each = length(groups)),
                       group = rep(groups, length(regions)),
                       coeff = rep(NA, length(regions) * length(groups)))
coeff <- c()
i = 1
for (i in 1:nrow(mu_coeff)){
  reg <- mu_coeff$region[i]
  group <- mu_coeff$group[i]
  if (reg == 1) {
    mu_int_reg <- mu_int
  } else {
    mu_int_reg <- sev_glm$mu.coefficients[reg] + mu_int
  }
  if (group == "major") {
    coeff[i] <- mu_int_reg
  } else {
    group_num <- match(group, groups) - 1
    coeff[i] <- mu_int + mu_groups[[group_num]]
  }
}

mu_coeff$coeff <- coeff

#simulate
nsim <- 10000
mu_coeff_sim <- t(sapply(1:nrow(mu_coeff), function(i)  rLOGNO(nsim, mu = mu_coeff$coeff[i], sigma = sev_glm$sigma.coefficients)))
mu_coeff_sim <- data.frame(mu_coeff_sim)
colnames(mu_coeff_sim) <- paste0("sim_", 1:ncol(mu_coeff_sim))

mu_coeff_sim
# Combine simulated values with original data frame
mu_coeff_combined <- cbind(mu_coeff, mu_coeff_sim)

########## Combine Severity and Frequency Models #########################
# Add additional costs (home and contents)
material_labour <- 0.35
contents <- 0.58
THC_2021 <- 1298835945
proportion <- 0.5
mu_coeff_sim_inf <- mu_coeff_sim*(1+contents)+ mu_coeff_sim*(1+contents+material_labour)*proportion
# Multiply severity and frequency simulations to obtain total cost
combined_sim <- mu_coeff_sim_inf*freq_sim
freq_sim
# Sum columns to obtain total cost for Storlysia across all regions and groups
combined_sim <- colSums(combined_sim)
# Import program reduction (hardcoded from excel 2020 figures)
reduction <- 1-0.38731162

# Reduce based on program
combined_sim_red <- (combined_sim +THC_2021)*reduction

# Import GDP data
gdp <- read.xlsx(file.path(data_dir, "Projection of Economic Costs FINAL.xlsx"),
                 sheet = "% of GDP - Voluntary",
                 rows = 8, cols = 2:53)
gdp_10percent <- gdp*0.1

# Determine confidence intervals
solvency_99.5 <- quantile(combined_sim_red, 0.995)
solvency_85 <- quantile(combined_sim_red, 0.85)
a<- solvency_99.5 - mean(combined_sim_red)
b <- solvency_85 - mean(combined_sim_red)
a
b
