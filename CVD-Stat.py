import numpy as np
import pandas as pd
import xlsxwriter
import datetime
import sys
import re
import os

import matplotlib
from pathlib import Path


path_Current_Directory = Path.cwd()
path_Data_Directory = path_Current_Directory / "Data"

path_OWID_Directory = path_Data_Directory / "OWID"
path_OWID_File = path_OWID_Directory / "owid.xlsx"

path_Output_DF = path_Data_Directory / "DataArray.xlsx"
str_Output_DF_SheetName = "DataArray"

path_Forecast_DF = path_Data_Directory / "Forecast.xlsx"
str_Forecast_DF_SheetName = "Forecast"

COUNTRIES_LIST = ["ZAF","GBR","FRA","DEU","ITA","POL","UKR","RUS"]
UKR_POPULATION = 43466822
POL_POPULATION = 37797000
RUS_POPULATION = 145912022



def main():

    print("\nCurrent Directory -> %s \n" % (path_Current_Directory))
    print("Data Directory -> %s \n" % (path_Data_Directory))
    print("OWID Directory -> %s \n" % (path_OWID_Directory))

    # Downloading a Data File from OWID
    answer = input("Would you like to update data from OWID site? Y/N > ")
    answer = answer.lower()
    if answer == "yes" or answer == "y":
        CVD_Download()
    
    # Load selected OWID data to a dataframe
    df_selected_data = pd.DataFrame()
    print("Reading a File with selected data to a DataFrame -> %s \n" % (path_Output_DF))
    df_selected_data = pd.read_excel((path_Output_DF), keep_default_na=False)


    # Prepare offset data for forecast
    UKR_offset = float(df_selected_data.loc[(df_selected_data["iso_code"] == "UKR") & (df_selected_data["date"] == "2022-01-10"), ["new_cases_smoothed"]].values[0])
    POL_offset = float(df_selected_data.loc[(df_selected_data["iso_code"] == "POL") & (df_selected_data["date"] == "2022-01-10"), ["new_cases_smoothed"]].values[0])
    RUS_offset = float(df_selected_data.loc[(df_selected_data["iso_code"] == "RUS") & (df_selected_data["date"] == "2022-01-10"), ["new_cases_smoothed"]].values[0])
    ZAF_offset = float(df_selected_data.loc[(df_selected_data["iso_code"] == "ZAF") & (df_selected_data["date"] == "2021-11-18"), ["new_cases_smoothed"]].values[0])
    
    # print(f"UKR offset: {UKR_offset}")
    # print(f"POL offset: {POL_offset}")
    # print(f"RUS offset: {RUS_offset}")
    # print(f"ZAF offset: {ZAF_offset}")


    # Select ZAF data
    df_ZAF_data = df_selected_data.loc[(df_selected_data["iso_code"] == "ZAF") & (df_selected_data["date"] >= "2021-11-18")]
    # print(df_ZAF_data.pivot(columns = ["iso_code"], values = ["date", "new_cases", "new_cases_smoothed", "new_cases_smoothed_per_million", "population"]))

    # Create UKR forecast - we are expecting start of a new wave at Jan.25   
    df_UKR_forecast = df_ZAF_data.copy()
    df_UKR_forecast["iso_code"] = "UKR-frcst"
    df_UKR_forecast["population"] = UKR_POPULATION
    df_UKR_forecast["continent"] = "Europe"
    df_UKR_forecast["location"] = "Ukraine"

    df_UKR_forecast["date"] += pd.DateOffset(days=68)
    df_UKR_forecast["year"] = df_UKR_forecast["date"].dt.year
    df_UKR_forecast["quarter"] = df_UKR_forecast["date"].dt.quarter
    df_UKR_forecast["month"] = df_UKR_forecast["date"].dt.month
    df_UKR_forecast["week"] = df_UKR_forecast["date"].dt.isocalendar().week
    df_UKR_forecast["day"] = df_UKR_forecast["date"].dt.day

    df_UKR_forecast["new_cases"] = (df_UKR_forecast["new_cases_per_million"] * df_UKR_forecast["population"] / 1000000) + UKR_offset - ZAF_offset
    df_UKR_forecast["new_cases"] = df_UKR_forecast["new_cases"].astype("int64")

    df_UKR_forecast["new_cases_smoothed"] = (df_UKR_forecast["new_cases_smoothed_per_million"] * df_UKR_forecast["population"] / 1000000) + UKR_offset - ZAF_offset
    #print(df_UKR_forecast.pivot(columns = ["iso_code"], values = ["date", "new_cases", "new_cases_smoothed", "new_cases_smoothed_per_million", "population"]))


    # Create POL forecast - we are expecting start of a new wave at Jan.11   
    df_POL_forecast = df_ZAF_data.copy()
    df_POL_forecast["iso_code"] = "POL-frcst"
    df_POL_forecast["population"] = POL_POPULATION
    df_POL_forecast["continent"] = "Europe"
    df_POL_forecast["location"] = "Poland"
 
    df_POL_forecast["date"] += pd.DateOffset(days=54)
    df_POL_forecast["year"] = df_POL_forecast["date"].dt.year
    df_POL_forecast["quarter"] = df_POL_forecast["date"].dt.quarter
    df_POL_forecast["month"] = df_POL_forecast["date"].dt.month
    df_POL_forecast["week"] = df_POL_forecast["date"].dt.isocalendar().week
    df_POL_forecast["day"] = df_POL_forecast["date"].dt.day

    df_POL_forecast["new_cases"] = (df_POL_forecast["new_cases_per_million"] * df_POL_forecast["population"] / 1000000) + POL_offset - ZAF_offset
    df_POL_forecast["new_cases"] = df_POL_forecast["new_cases"].astype("int64")

    df_POL_forecast["new_cases_smoothed"] = (df_POL_forecast["new_cases_smoothed_per_million"] * df_POL_forecast["population"] / 1000000) + POL_offset - ZAF_offset
    #print(df_POL_forecast.pivot(columns = ["iso_code"], values = ["date", "new_cases", "new_cases_smoothed", "new_cases_smoothed_per_million", "population"]))


    # Create RUS forecast - we are expecting start of a new wave at Jan.25  
    df_RUS_forecast = df_ZAF_data.copy()
    df_RUS_forecast["iso_code"] = "RUS-frcst"
    df_RUS_forecast["population"] = RUS_POPULATION
    df_RUS_forecast["continent"] = "Europe"
    df_RUS_forecast["location"] = "Russia"
 
    df_RUS_forecast["date"] += pd.DateOffset(days=68)
    df_RUS_forecast["year"] = df_RUS_forecast["date"].dt.year
    df_RUS_forecast["quarter"] = df_RUS_forecast["date"].dt.quarter
    df_RUS_forecast["month"] = df_RUS_forecast["date"].dt.month
    df_RUS_forecast["week"] = df_RUS_forecast["date"].dt.isocalendar().week
    df_RUS_forecast["day"] = df_RUS_forecast["date"].dt.day

    df_RUS_forecast["new_cases"] = (df_RUS_forecast["new_cases_per_million"] * df_RUS_forecast["population"] / 1000000) + RUS_offset - ZAF_offset
    df_RUS_forecast["new_cases"] = df_RUS_forecast["new_cases"].astype("int64")

    df_RUS_forecast["new_cases_smoothed"] = (df_RUS_forecast["new_cases_smoothed_per_million"] * df_RUS_forecast["population"] / 1000000) + RUS_offset - ZAF_offset
    #print(df_RUS_forecast.pivot(columns = ["iso_code"], values = ["date", "new_cases", "new_cases_smoothed", "new_cases_smoothed_per_million", "population"]))


    # Add Forecast Data to Selected Data
    df_selected_data = df_selected_data.append(df_UKR_forecast, sort=False)
    df_selected_data = df_selected_data.append(df_POL_forecast, sort=False)
    df_selected_data = df_selected_data.append(df_RUS_forecast, sort=False)


    # Save Forecast Data
    print("Writing Forecast Data to a Directory -> %s \n" % (path_Data_Directory))
    writer = pd.ExcelWriter(path_Forecast_DF, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    df_selected_data.to_excel(writer, sheet_name=str_Forecast_DF_SheetName, index=False)
    writer.save()



def CVD_Download():

    os.system("curl -f -o %(filename)s https://covid.ourworldindata.org/data/owid-covid-data.xlsx" % {
        "filename": path_OWID_File })

    print("A data from OWID site downloaded to -> %s \n" % (path_OWID_File))


    # Load OWID data to a dataframe
    df_data_array = pd.DataFrame()
    print("Reading a File to a DataFrame -> %s \n" % (path_OWID_File))
    df_data_array = pd.read_excel((path_OWID_File), keep_default_na=False)


    # Drop Excess Columns
    df_data_array = df_data_array.drop(columns=["reproduction_rate", "icu_patients", "icu_patients_per_million", "hosp_patients", "hosp_patients_per_million",	
            "weekly_icu_admissions", "weekly_icu_admissions_per_million", "weekly_hosp_admissions", "weekly_hosp_admissions_per_million", "new_tests",	
            "total_tests", "total_tests_per_thousand", "new_tests_per_thousand", "new_tests_smoothed", "new_tests_smoothed_per_thousand", "positive_rate",	
            "tests_per_case", "tests_units", "total_vaccinations", "people_vaccinated", "people_fully_vaccinated",	"total_boosters", "new_vaccinations",	
            "new_vaccinations_smoothed", "total_vaccinations_per_hundred", "people_vaccinated_per_hundred",	"people_fully_vaccinated_per_hundred",
            "total_boosters_per_hundred", "new_vaccinations_smoothed_per_million", "new_people_vaccinated_smoothed", "new_people_vaccinated_smoothed_per_hundred",
            "stringency_index",	"median_age", "aged_65_older", "aged_70_older",	"gdp_per_capita", "extreme_poverty", "cardiovasc_death_rate", "diabetes_prevalence",
            "female_smokers", "male_smokers", "handwashing_facilities", "hospital_beds_per_thousand", "life_expectancy", "human_development_index"])


    # Define Data Types
    df_data_array["date"] = pd.to_datetime(df_data_array["date"], format="%Y-%m-%d")
    
    df_data_array["total_cases"] = df_data_array["total_cases"].fillna(0)
    df_data_array.loc[df_data_array["total_cases"] == "", ["total_cases"]] = 0
    df_data_array["total_cases"] = df_data_array["total_cases"].astype("int64")
    
    df_data_array["new_cases"] = df_data_array["new_cases"].fillna(0)
    df_data_array.loc[df_data_array["new_cases"] == "", ["new_cases"]] = 0
    df_data_array["new_cases"] = df_data_array["new_cases"].astype("int64")
    
    df_data_array["new_cases_smoothed"] = df_data_array["new_cases_smoothed"].fillna(0)
    df_data_array.loc[df_data_array["new_cases_smoothed"] == "", ["new_cases_smoothed"]] = 0
    df_data_array["new_cases_smoothed"] = df_data_array["new_cases_smoothed"].astype("float64")

    df_data_array["total_cases_per_million"] = df_data_array["total_cases_per_million"].fillna(0)
    df_data_array.loc[df_data_array["total_cases_per_million"] == "", ["total_cases_per_million"]] = 0
    df_data_array["total_cases_per_million"] = df_data_array["total_cases_per_million"].astype("float64")

    df_data_array["new_cases_per_million"] = df_data_array["new_cases_per_million"].fillna(0)
    df_data_array.loc[df_data_array["new_cases_per_million"] == "", ["new_cases_per_million"]] = 0
    df_data_array["new_cases_per_million"] = df_data_array["new_cases_per_million"].astype("float64")	
    
    df_data_array["new_cases_smoothed_per_million"] = df_data_array["new_cases_smoothed_per_million"].fillna(0)
    df_data_array.loc[df_data_array["new_cases_smoothed_per_million"] == "", ["new_cases_smoothed_per_million"]] = 0
    df_data_array["new_cases_smoothed_per_million"] = df_data_array["new_cases_smoothed_per_million"].astype("float64")
     
    df_data_array["total_deaths"] = df_data_array["total_deaths"].fillna(0)
    df_data_array.loc[df_data_array["total_deaths"] == "", ["total_deaths"]] = 0
    df_data_array["total_deaths"] = df_data_array["total_deaths"].astype("int64")

    df_data_array["new_deaths"] = df_data_array["new_deaths"].fillna(0)
    df_data_array.loc[df_data_array["new_deaths"] == "", ["new_deaths"]] = 0
    df_data_array["new_deaths"] = df_data_array["new_deaths"].astype("int64")

    df_data_array["new_deaths_smoothed"] = df_data_array["new_deaths_smoothed"].fillna(0)
    df_data_array.loc[df_data_array["new_deaths_smoothed"] == "", ["new_deaths_smoothed"]] = 0
    df_data_array["new_deaths_smoothed"] = df_data_array["new_deaths_smoothed"].astype("float64")
    
    df_data_array["total_deaths_per_million"] = df_data_array["total_deaths_per_million"].fillna(0)
    df_data_array.loc[df_data_array["total_deaths_per_million"] == "", ["total_deaths_per_million"]] = 0
    df_data_array["total_deaths_per_million"] = df_data_array["total_deaths_per_million"].astype("float64")

    df_data_array["new_deaths_per_million"] = df_data_array["new_deaths_per_million"].fillna(0)
    df_data_array.loc[df_data_array["new_deaths_per_million"] == "", ["new_deaths_per_million"]] = 0
    df_data_array["new_deaths_per_million"] = df_data_array["new_deaths_per_million"].astype("float64")

    df_data_array["new_deaths_smoothed_per_million"] = df_data_array["new_deaths_smoothed_per_million"].fillna(0)
    df_data_array.loc[df_data_array["new_deaths_smoothed_per_million"] == "", ["new_deaths_smoothed_per_million"]] = 0
    df_data_array["new_deaths_smoothed_per_million"] = df_data_array["new_deaths_smoothed_per_million"].astype("float64")
    
    df_data_array["population"] = df_data_array["population"].fillna(0)
    df_data_array.loc[df_data_array["population"] == "", ["population"]] = 0
    df_data_array["population"] = df_data_array["population"].astype("int64")

    df_data_array["population_density"] = df_data_array["population_density"].fillna(0)
    df_data_array.loc[df_data_array["population_density"] == "", ["population_density"]] = 0
    df_data_array["population_density"] = df_data_array["population_density"].astype("float64")
    
    df_data_array["excess_mortality_cumulative_absolute"] = df_data_array["excess_mortality_cumulative_absolute"].fillna(0)
    df_data_array.loc[df_data_array["excess_mortality_cumulative_absolute"] == "", ["excess_mortality_cumulative_absolute"]] = 0
    df_data_array["excess_mortality_cumulative_absolute"] = df_data_array["excess_mortality_cumulative_absolute"].astype("int64")

    df_data_array["excess_mortality_cumulative"] = df_data_array["excess_mortality_cumulative"].fillna(0)
    df_data_array.loc[df_data_array["excess_mortality_cumulative"] == "", ["excess_mortality_cumulative"]] = 0
    df_data_array["excess_mortality_cumulative"] = df_data_array["excess_mortality_cumulative"].astype("int64")

    df_data_array["excess_mortality"] = df_data_array["excess_mortality"].fillna(0)
    df_data_array.loc[df_data_array["excess_mortality"] == "", ["excess_mortality"]] = 0
    df_data_array["excess_mortality"] = df_data_array["excess_mortality"].astype("int64")

    df_data_array["excess_mortality_cumulative_per_million"] = df_data_array["excess_mortality_cumulative_per_million"].fillna(0)
    df_data_array.loc[df_data_array["excess_mortality_cumulative_per_million"] == "", ["excess_mortality_cumulative_per_million"]] = 0
    df_data_array["excess_mortality_cumulative_per_million"] = df_data_array["excess_mortality_cumulative_per_million"].astype("float64")
    
    df_data_array.insert(0, "day", 0)
    df_data_array.insert(0, "week", 0)
    df_data_array.insert(0, "month", 0)    
    df_data_array.insert(0, "quarter", 0)
    df_data_array.insert(0, "year", 0)

    df_data_array["year"] = df_data_array["date"].dt.year
    df_data_array["quarter"] = df_data_array["date"].dt.quarter
    df_data_array["month"] = df_data_array["date"].dt.month
    df_data_array["week"] = df_data_array["date"].dt.isocalendar().week
    df_data_array["day"] = df_data_array["date"].dt.day
    

    # Slice data for South Africa, UK, France, Germany, Italy, Poland, Ukraine, Russia
    df_short_array = df_data_array.loc[df_data_array["iso_code"].isin(COUNTRIES_LIST)]


    # Save Short DataFrame
    print("Writing the Data to a Directory -> %s \n" % (path_Data_Directory))
    writer = pd.ExcelWriter(path_Output_DF, engine='xlsxwriter') # pylint: disable=abstract-class-instantiated
    df_short_array.to_excel(writer, sheet_name=str_Output_DF_SheetName, index=False)
    writer.save()



if __name__ == "__main__":
    main()
