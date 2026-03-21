📊 HR Attrition Analytics & Automation System

<img width="1897" height="831" alt="HR Attrition Dashboard" src="https://github.com/user-attachments/assets/873fe496-b55b-46f6-bec9-d2907683109c" />

📌 Project Overview
This project is a comprehensive End-to-End HR Analytics Solution designed to identify key drivers of employee turnover. It transforms raw workforce data into a dynamic, interactive dashboard that helps HR leaders move from reactive reporting to proactive retention strategies.
The system analyzes a dataset of 2,925 employees, focusing on performance, demographics, and operational strain.
🚀 Key Features

    Automated Data Pipeline: Uses Power Query (M Language) to clean, transform, and bucket raw data (Age, Distance, Satisfaction) without manual intervention.
    Interactive VBA Dashboard: Custom VBA macros to synchronize checkboxes with Slicers for a seamless UI/UX.
    Dynamic KPI Tracking: Real-time calculation of Attrition Rates, Satisfaction Levels, and Performance Ratios.
    Risk Heatmapping: Visual identification of high-risk segments (e.g., "Very-far" commuters and "High-performing" leavers).

🛠️ Technical Implementation
1. Data Transformation (Power Query)
The data was processed using M-code to create logical buckets:
powerquery

#"Added Conditional Column1" = Table.AddColumn(#"Removed Columns", "Age Bucket", each 
    if [Age] <= 25 then "18-25" 
    else if [Age] <= 35 then "26-35" 
    else if [Age] <= 45 then "36-45" 
    else "56 Plus")

2. Automation (VBA)
Custom scripts were written to handle multi-pivot filtering through a single interface:
vba

Sub FilterMultiplePivots()
    ' Syncs Checkboxes with Slicer Items
    sc.SlicerItems("Far").Selected = (ActiveSheet.CheckBoxes("Check Box 1").Value = 1)
    ' ...
End Sub

📈 Insights & Findings

    High Performer Loss: 84% of total attrition (413/492) consists of High Performers.
    Location Sensitivity: Employees in the "Very-far" category have a 23.38% attrition rate.
    Overtime Correlation: 57% of those who left were required to work Overtime.

📂 Repository Structure

    /Data: Raw CSV/Excel employee datasets.
    /Dashboard: The final .xlsm file with the interactive UI.
    /Scripts: Separate .bas files for the VBA modules.
    /Reports: Executive Summary and PDF exports.

👨‍💻 How to Use

    Clone the repository.
    Open the .xlsm file.
    Click the "Refresh All" button (via Macro) to sync data.
    Use the Checkboxes to filter data by distance and role.
