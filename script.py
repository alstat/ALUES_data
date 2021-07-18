import os
import numpy as np
import pandas as pd

FILENAME = "CropInfo.xls"
OUTFOLDER = "CSV"
RDOCS = "RDOCS"

def crop_names(file):
    xl = pd.ExcelFile(file)
    
    crops = []
    for i in xl.sheet_names:
        if (i != "About") and (i != "FactorList") and (i != "Form") and (i != "Information"):
            crops.append(i)
    
    return crops

def export_csvs(file, folder):
    crops_sheets = crop_names(file)

    trrn_start  = 12
    soil_start  = 35
    water_start = 58
    temp_start  = 81
    
    for sheet in crops_sheets:
        sheet_df = pd.read_excel(file, sheet_name=sheet)

        trrn_nrow  = sheet_df.iloc[5, 2]
        soil_nrow  = sheet_df.iloc[6, 2]
        water_nrow = sheet_df.iloc[7, 2]
        temp_nrow  = sheet_df.iloc[8, 2]

        terrain_df = sheet_df.iloc[trrn_start:(trrn_start+trrn_nrow+1), 1:9]
        terrain_df.columns = terrain_df.iloc[0, :]
        terrain_df = terrain_df.iloc[1:,:].set_index("Code")
        terrain_df.to_csv(os.path.join(folder, sheet.upper() + "Terrain.csv"))

        soil_df = sheet_df.iloc[soil_start:(soil_start+soil_nrow+1), 1:9]
        soil_df.columns = soil_df.iloc[0, :]
        soil_df = soil_df.iloc[1:,:].set_index("Code")
        soil_df.to_csv(os.path.join(folder, sheet.upper() + "Soil.csv"))

        water_df = sheet_df.iloc[water_start:(water_start+water_nrow+1), 1:9]
        water_df.columns = water_df.iloc[0, :]
        water_df = water_df.iloc[1:,:].set_index("Code")
        water_df.to_csv(os.path.join(folder, sheet.upper() + "Water.csv"))

        temp_df = sheet_df.iloc[temp_start:(temp_start+temp_nrow+1), 1:9]
        temp_df.columns = temp_df.iloc[0, :]
        temp_df = temp_df.iloc[1:,:].set_index("Code")
        temp_df.to_csv(os.path.join(folder, sheet.upper() + "Temp.csv"))

export_csvs(FILENAME, OUTFOLDER)

def export_docs(file, folder):
    crops_sheets = crop_names(file)
    for sheet in crops_sheets:
        sheet_df = pd.read_excel(file, sheet_name=sheet)
            
        trrn_nrow  = sheet_df.iloc[5, 2]
        soil_nrow  = sheet_df.iloc[6, 2]
        water_nrow = sheet_df.iloc[7, 2]
        temp_nrow  = sheet_df.iloc[8, 2]
        
        trrn_start  = 12
        soil_start  = 35
        water_start = 58
        temp_start  = 81

        trrn_label  = sheet_df.iloc[trrn_start-1, 0].strip().lower()
        soil_label  = sheet_df.iloc[soil_start-1, 0].strip().lower()[:4]
        water_label = sheet_df.iloc[water_start-1, 0].strip().lower()
        temp_label  = sheet_df.iloc[temp_start-1, 0].strip().lower()[:4]
        
        characteristics_labels = [trrn_label, soil_label, water_label, temp_label]
        characteristics_starts = [trrn_start, soil_start, water_start, temp_start]
        characteristics_nrows = [trrn_nrow, soil_nrow, water_nrow, temp_nrow]

        for j in range(len(characteristics_labels)):
            descriptions = sheet_df.iloc[(characteristics_starts[j]+1):(characteristics_starts[j]+characteristics_nrows[j]+1), 9].values
            desc = []
            for d in descriptions:
                try:
                    desc.append(d.replace("%", "percent"))
                except:
                    desc.append("NA")
            codes = sheet_df.iloc[(characteristics_starts[j]+1):(characteristics_starts[j]+characteristics_nrows[j]+1), 1].values
            
            generate_doc(sheet, characteristics_labels[j], codes, descriptions, folder)

def generate_doc(crop_name, characteristics, codes, descriptions, outfolder):

    header_params = {"crop_name": crop_name, "characteristics": characteristics}
    header = """#' %(crop_name)s %(characteristics)s requirement for land evaluation
    #' 
    #' A dataset containing the %(characteristics)s characteristics of the crop requirements for farming Banana.
    #' 
    #' @details 
    #' The following are the factors for evaluation: \n#'
    """ % header_params

    items = """#' \itemize{\n#'"""
    for (i,j) in zip(codes, descriptions):
        items = items + " \item " + str(i) + " - " + str(j) + "\n#'"
    items += " }"
    items = items.replace("%", "\%")
    footer = """
    #' @seealso 
    #' \itemize{
    #'  \item Yen, B. T., Pheng, K. S., and Hoanh, C. T. (2006). \emph{LUSET: Land Use Suitability Evaluation Tool User's Guide}. International Rice Research Institute.
    #'  }
    #' 
    #' @docType data
    #' @keywords dataset
    """

    dataformat = "#' @format A data frame with " + str(codes.shape[0]) + " rows and 8 columns"
    name = "\n#' @name " + crop_name.upper() + characteristics.title() + "\nNULL"

    doc = header + items + footer + dataformat + name
    
    with open(outfolder + "/" + crop_name.upper() + characteristics.title() + ".R", "w") as text_file:
        text_file.write(doc)
    return doc

export_docs(FILENAME, RDOCS)