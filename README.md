# Advanced MATLAB Automation Scripts for LightTools

This repository contains a collection of powerful and robust MATLAB scripts designed to automate complex analysis and visualization tasks within Synopsys LightTools. These scripts leverage the LightTools COM API through its `.NET` interface to provide functionality that goes beyond simple command execution, demonstrating advanced techniques for data extraction, interactive filtering, and dynamic visualization.

## Features

This collection showcases two primary automation workflows:

### 1. Comprehensive Model Interrogation (`InterrogateLightToolsModel.m`)

This script acts as a powerful "discovery" tool, allowing you to programmatically dump all known properties of any specified object in a LightTools model.

*   **Connects to Specific Session:** Uses a Process ID (PID) to connect to a specific running instance of LightTools.
*   **Processes Multiple Keys:** Takes a list of LightTools data access keys as input, enabling bulk analysis of many different objects in a single run.
*   **Dynamic Property Discovery:** Automatically uses the `DbKeyDump` API function to learn all available properties for an object, eliminating the need to hardcode property names.
*   **Handles Scalar and Array Data:**
    *   Retrieves standard scalar properties (names, positions, numeric values) using `DbGet`.
    *   Intelligently identifies 2D array data (like illumination or intensity meshes) and uses the specialized `GetMeshData` function to extract the full numerical matrices.
*   **Dual Output Format:**
    *   **`.csv` file:** A human-readable summary of all scalar properties for quick review and analysis in spreadsheet software.
    *   **`.mat` file:** A complete MATLAB data file that stores both the scalar results and the full numerical arrays from meshes, preserving all data for further programmatic analysis.

### 2. Interactive Ray Path Visualization (`VisualizeTopRayPaths.m`)

This is an advanced analysis script that helps users identify and visualize the most significant ray paths in an optical system based on their power contribution.

*   **Interactive Filtering:** Instead of requiring users to hardcode names, the script:
    1.  Retrieves all ray path data from a simulation.
    2.  Dynamically finds all unique source names and final surface names.
    3.  Presents these names to the user in interactive list dialogs for easy and error-free filtering.
*   **Power Percentage Interval Analysis:**
    *   Prompts the user to define one or more power percentage intervals, such as `[[100, 70], [70, 30]]`, to analyze different power "bands" of rays.
*   **Automated Visibility Control:**
    *   For each defined interval, it calculates which rays fall within that power bracket based on the total power of the filtered set.
    *   It programmatically sets the visibility of all other rays to "off" and the selected rays to "on" using the `RayPathVisibleAt` property.
*   **Automated Screenshot Capture:**
    *   After setting the visibility for each interval, it automatically activates the 3D Design View, ensures the display is updated (`RecalcNow`), and captures a high-quality screenshot using the Windows clipboard and Java AWT integration.
    *   Each screenshot is saved with a descriptive, unique filename that includes the interval and filter criteria.
*   **Data Archiving:**
    *   Along with the screenshot, it saves a `.mat` file containing the detailed data of the specific rays that were made visible for that interval, allowing for further quantitative analysis.
*   **Performance Optimization:** Employs `DBUpdateOff` and `RecalcOff` during bulk visibility changes to prevent the LightTools UI from redrawing after every single command, dramatically speeding up the process from minutes to seconds.
*   **State Restoration:** After the script completes, it restores the visibility of all ray paths in the LightTools model.

## Requirements

*   **MATLAB:** A modern version with `.NET` interoperability support.
*   **LightTools:** A version compatible with the API calls (scripts developed and tested with LightTools 8.4.0). The LightTools COM server must be registered during its installation.
*   **Operating System:** Windows.
*   **A Running LightTools Session:** These scripts are designed to connect to an active LightTools instance. You must have your model loaded and the relevant simulation run *before* executing the MATLAB script.

## How to Use

1.  **Prepare LightTools:**
    *   Launch LightTools and load your model.
    *   Run any necessary simulations to generate the data you want to analyze (e.g., a forward simulation with "Save Ray Data" and "Collect Ray Paths" enabled).
    *   Find the **Process ID (PID)** of your LightTools session, which is displayed in the LightTools Console window upon startup.

2.  **Configure the MATLAB Script:**
    *   Open either `InterrogateLightToolsModel.m` or `VisualizeTopRayPaths.m` in MATLAB.
    *   Navigate to the **Configuration** section at the top of the script.
    *   Update the `lightToolsVersion` variable to match your installed version.
    *   Update the `pid` variable with the correct Process ID from your running LightTools session.
    *   For `InterrogateLightToolsModel.m`, edit the `keysToProcess` cell array to list the data access strings of the objects you want to inspect.
    *   For `VisualizeTopRayPaths.m`, ensure the `baseKeyForRayPaths` points to the correct `FORWARD_SIM_FUNCTION` in your model.

3.  **Run the Script:**
    *   Execute the script from the MATLAB editor or command window.
    *   Follow any on-screen prompts (like the interval and filter dialogs in the visualization script).

4.  **Check Your Output:**
    *   Navigate to the `saveDirectory` you specified. You will find the generated `.mat`, `.csv`, or `.png` files with descriptive, timestamped names.

---
