import pandas as pd

# Main production dataframe
# Each row represents a unique product/line speed combination

data = {
    "Product": [
        "Alpha Bar", "Alpha Bar",
        "Beta Snack", "Beta Snack",
        "Gamma Wrap", "Gamma Wrap"
    ],
    "Line_Speed": ["Slow", "Fast", "Slow", "Fast", "Slow", "Fast"],

    # Bill of Materials
    "Primary_Material": ["Film-A", "Film-A", "Film-B", "Film-B", "Film-C", "Film-C"],
    "Primary_Material_Spec": ["2.5mil", "2.5mil", "3.0mil", "3.0mil", "1.8mil", "1.8mil"],
    "Adhesive": ["Adh-1", "Adh-1", "Adh-2", "Adh-2", "Adh-1", "Adh-1"],
    "Adhesive_Rate_gpm": [12.5, 14.0, 11.0, 13.5, 10.0, 12.0],

    # Machine Parameters
    "Temperature_F": [325, 340, 310, 330, 300, 315],
    "Pressure_psi": [45, 50, 42, 48, 40, 45],
    "Line_Speed_fpm": [150, 250, 140, 230, 130, 220],

    # Process Times
    "Setup_Time_min": [30, 30, 35, 35, 25, 25],
    "Changeover_Time_min": [45, 45, 50, 50, 40, 40],
    "Units_Per_Hour": [1200, 2000, 1100, 1800, 1000, 1700],
    "Units_Per_Shift": [9600, 16000, 8800, 14400, 8000, 13600],

    # Quality Check Parameters
    "Adhesion_Min": [85, 85, 80, 80, 90, 90],
    "Adhesion_Max": [95, 95, 92, 92, 98, 98],
    "Thickness_Min_mil": [2.3, 2.3, 2.8, 2.8, 1.6, 1.6],
    "Thickness_Max_mil": [2.7, 2.7, 3.2, 3.2, 2.0, 2.0],

    # Post Processing
    "Post_Processing": ["Slitting", "Slitting", "None", "None", "Coating", "Coating"],
    "Post_Process_Notes": [
        "Slit to 6in rolls", "Slit to 6in rolls",
        "N/A", "N/A",
        "Apply top coat after lamination", "Apply top coat after lamination"
    ]
}

df_production = pd.DataFrame(data)

# Label dataframe
label_data = {
    "Product": ["Alpha Bar", "Beta Snack", "Gamma Wrap"],
    "Customer": ["Acme Foods", "Beta Brands", "Gamma Co"],
    "Finished_Width_in": [6.0, 8.0, 4.5],
    "Finished_Length_in": [12.0, 10.0, 8.0],
    "Roll_Weight_lbs": [25.0, 30.0, 18.0],
    "Label_1": ["Primary ID", "Primary ID", "Primary ID"],
    "Label_2": ["Lot Code", "Lot Code", "Lot Code"],
    "Label_3": ["Customer Spec", "Customer Spec", "Customer Spec"],
    "Label_4": ["Shipping", "Shipping", "Shipping"],
    "Hazmat": [False, False, False],
    "Special_Instructions": [
        "Keep dry", "Temperature sensitive", "None"
    ]
}

df_labels = pd.DataFrame(label_data)

if __name__ == "__main__":
    print("Production DataFrame:")
    print(df_production.to_string())
    print("\nLabel DataFrame:")
    print(df_labels.to_string())