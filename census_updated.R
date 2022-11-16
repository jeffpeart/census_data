library(httr)
library(jsonlite)
library(dplyr)
library(stringr)
library(tidyr)
library(tidyverse)
library(openxlsx)

### 2000 (Decennial Survey)

desired_vars = c(totalpriv_male = "P051004",
                 privcomp_male = "P051005",
                 seinc_male = "P051006",
                 seuninc_male ="P051011",
                 totalpriv_fem = "P051036",
                 privcomp_fem = "P051037",
                 seinc_fem = "P051038",
                 seuninc_fem = "P051043")
emp_2000 = get_decennial(
  geography = "state",
  variables = desired_vars,
  year = 2000
)
emp_2000['year'] = 2000

colnames(emp_2000)[4] ="estimate"

### 2010 - 2015
desired_vars = c(total_priv = "S2408_C01_002",
                 total_privcomp = "S2408_C01_003",
                 total_seinc = "S2408_C01_004",
                 total_seuninc = "S2408_C01_009",
                 totalpriv_male = "S2408_C02_002",
                 privcomp_male = "S2408_C02_003",
                 seinc_male = "S2408_C02_004",
                 seuninc_male ="S2408_C02_009",
                 totalpriv_fem = "S2408_C04_002",
                 privcomp_fem = "S2408_C04_003",
                 seinc_fem = "S2408_C04_004",
                 seuninc_fem = "S2408_C04_009")
x = list()

for (i in (c(2010:2014))){
  emp = get_acs(
    geography = "state",
    variables = desired_vars,
    year = i
    )

  emp = emp %>% pivot_wider(id_cols = c(NAME,GEOID), 
                            names_from = variable, 
                            values_from = estimate)

  emp$totalpriv_fem = ((1-(emp$totalpriv_male/100)))* emp$total_priv
  emp$privcomp_fem = ((1-(emp$privcomp_male/100)))* emp$total_privcomp
  emp$seinc_fem = ((1-(emp$seinc_male/100)))* emp$total_seinc
  emp$seuninc_fem = ((1-(emp$seuninc_male/100)))* emp$total_seuninc
  
  emp$totalpriv_male = (emp$totalpriv_male/100) * emp$total_priv
  emp$privcomp_male = (emp$privcomp_male/100) * emp$total_privcomp
  emp$seinc_male = (emp$seinc_male/100) * emp$total_seinc
  emp$seuninc_male = (emp$seuninc_male/100) * emp$total_seuninc
  
  emp = emp %>% pivot_longer(cols = c(totalpriv_male,privcomp_male,
                                      seinc_male, seuninc_male,
                                      totalpriv_fem, privcomp_fem,
                                      seinc_fem, seuninc_fem), 
                             names_to = 'variable',
                             values_to = 'estimate')
  emp$year = i
  
  emp = emp[c('NAME','year','GEOID','estimate','variable')]
  
  x[[i]] = emp
}

### 2015 - 2020 (American Community Survey)
  
desired_vars = c(totalpriv_male = "S2408_C02_002",
                 privcomp_male = "S2408_C02_003",
                 seinc_male = "S2408_C02_004",
                 seuninc_male ="S2408_C02_009",
                 totalpriv_fem = "S2408_C04_002",
                 privcomp_fem = "S2408_C04_003",
                 seinc_fem = "S2408_C04_004",
                 seuninc_fem = "S2408_C04_009")

for (i in (c(2015:2020))){
  
  emp = get_acs(
    geography = "state",
    variables = desired_vars,
    year = i
  )
  emp['year'] = i
  emp = subset(emp, select = -c(moe) )
  x[[i]] = emp
}

### 2021 (American Community Survey 1 year)

emp = get_acs(
  geography = "state",
  variables = desired_vars,
  survey = "acs1",
  year = 2021
)
emp['year'] = 2021
emp = subset(emp, select = -c(moe) )
x[[2021]] = emp

## merge everything

merged = do.call(rbind, x)

final = rbind(merged, emp_2000)
final = final[order(final$year),]
final2 = final
final2$gender = word(final2$variable, 2, sep = fixed("_"))
final2$variable = word(final2$variable, 1, sep= fixed("_"))


final2 = final2 %>% pivot_wider(id_cols = c(NAME,GEOID,year,gender),
                      names_from = variable,
                      values_from = estimate)

final = final %>% pivot_wider(id_cols = c(NAME,GEOID,year),
                               names_from = variable,
                               values_from = estimate)

attach(final)

final$perc_totalpriv_male = (totalpriv_male / (totalpriv_male + totalpriv_fem))
final$perc_privcomp_male = (privcomp_male / (privcomp_male + privcomp_fem))
final$perc_seinc_male = (seinc_male / (seinc_male + seinc_fem))
final$perc_seuninc_male = (seuninc_male / (seuninc_male + seuninc_fem))
final$perc_totalpriv_fem = (totalpriv_fem / (totalpriv_male + totalpriv_fem))
final$perc_privcomp_fem = (privcomp_fem / (privcomp_male + privcomp_fem))
final$perc_seinc_fem = (seinc_fem/ (seinc_male + seinc_fem))
final$perc_seuninc_fem = (seuninc_fem / (seuninc_male + seuninc_fem))

final = final %>% pivot_longer(cols = colnames(final[4:length(colnames(final))]),
                               names_to = 'variable',
                               values_to = 'estimate')

### Create Excel Workbook and add metadata

wb <- createWorkbook()
addWorksheet(wb, "metadata")
writeData(wb, sheet = "metadata", x = final2)

## add state worksheets

states = as.list(unique(final['NAME']))[[1]][1:52]
for (i in 1:length(states)){
  addWorksheet(wb, states[i])
  export = final[final['NAME'] == 
                    states[i],][c('estimate','variable','year')] %>%
    pivot_wider(
      names_from = year, values_from = estimate
    )
  writeData(wb, sheet = states[i], x = export)
}

saveWorkbook(wb, 'time_series1.xlsx')

##### Business Characteristics #####

bs = list()
for (i in (c(2017:2020))){
  res = GET(sprintf("https://api.census.gov/data/%s/abscs?get=NAME,GEO_ID,NAICS2017_LABEL,SEX,SEX_LABEL,YEAR,EMPSZFI,EMPSZFI_LABEL,FIRMPDEMP,EMP&for=state:*&key=1c8d16edd6d5c9c7dfa9027309dd6bd95802d614",i))
  #res = GET("https://api.census.gov/data/2020/abscs?get=NAME,GEO_ID,YEAR,NAICS2017_LABEL,SEX,SEX_LABEL,EMPSZFI,EMPSZFI_LABEL,FIRMPDEMP,EMP&for=state:*&key=1c8d16edd6d5c9c7dfa9027309dd6bd95802d614")
  df= as.data.frame(fromJSON(rawToChar(res$content)))
  colnames(df)=df[1,]
  df = df[-1,]
  df = df[c('NAME','SEX_LABEL','EMPSZFI_LABEL','FIRMPDEMP','EMP',"YEAR")]
  df = df %>% rename("state" = NAME, "gender" = SEX_LABEL, "firm_size" = EMPSZFI_LABEL,
                "firms_quantity" = FIRMPDEMP,"employee_quantity" = EMP, "year" = YEAR)
  bs[[i]] = df
}

merged = do.call(rbind, bs)
merged$firms_quantity= as.numeric(merged$firms_quantity)
merged$employee_quantity= as.numeric(merged$employee_quantity)
merged$year = as.numeric(merged$year)
merged = merged[order(merged$state, merged$gender,decreasing = TRUE),]


merged2 = merged %>% 
  pivot_longer(cols = c(firms_quantity,employee_quantity),
               names_to = 'variable',
               values_to = 'estimate')

### Create Excel File ###

bsx <- createWorkbook()
addWorksheet(bsx, "metadata")

writeData(bsx, sheet = "metadata", x = merged)

states = as.list(unique(merged['state']))[[1]]
for (i in 1:length(states)){
  addWorksheet(bsx, states[i])
  export = merged2[merged2['state'] == 
                   states[i],] %>%
    pivot_wider(
      names_from = year, 
      values_from = estimate
    )
  export = export[order(export$gender,decreasing = TRUE),]
  writeData(bsx, sheet = states[i], x = export)
}

saveWorkbook(bsx, 'business_stats.xlsx')


